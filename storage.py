"""
storage.py — banco de preços da empresa (SQLAlchemy).

Sem DATABASE_URL  -> SQLite local (apenas para testes; na nuvem some ao reiniciar).
Com DATABASE_URL  -> qualquer Postgres (Supabase, Neon, RDS...).
"""
from __future__ import annotations

import hashlib
import os
import uuid
from datetime import datetime, timedelta, timezone

import pandas as pd
from sqlalchemy import (Column, Date, DateTime, Float, Integer, MetaData, String, Table, Text,
                        create_engine, delete, func, insert, select)

from core import normalizar

BRT = timezone(timedelta(hours=-3))  # horário de Brasília (sem horário de verão)
SIT_NOVO = "Novo"
SIT_EXISTE = "Já no banco"
SIT_REPETIDO = "Repetido na planilha"

metadata = MetaData()
precos = Table(
    "precos", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("codigo", String(100)),
    Column("descricao", Text, nullable=False),
    Column("descricao_norm", Text, nullable=False),          # chave de busca (já normalizada)
    Column("valor_unitario", Float, nullable=False),
    Column("fornecedor", String(255)),
    Column("vendedor", String(255)),
    Column("data_cotacao", Date),
    Column("origem", String(255)),                           # de qual licitação/planilha veio
    Column("lote", String(40), index=True),                  # identifica uma importação (para desfazer)
    Column("criado_em", DateTime, nullable=False),
    Column("hash_linha", String(40), nullable=False, unique=True),  # evita duplicar o mesmo registro
)


# ----------------------------------------------------------------------------
# Conexão
# ----------------------------------------------------------------------------
def _ajustar_url(url: str) -> str:
    """Aceita a string como o Supabase/Neon entregam (postgres:// ou postgresql://)."""
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    if url.startswith("postgresql://"):
        url = "postgresql+psycopg2://" + url[len("postgresql://"):]
    return url


def criar_engine(url: str | None = None):
    if not url:
        os.makedirs("dados_locais", exist_ok=True)
        url = "sqlite:///dados_locais/banco_precos.db"
    return create_engine(_ajustar_url(url), pool_pre_ping=True)


def eh_remoto(engine) -> bool:
    return engine.dialect.name != "sqlite"


def inicializar(engine) -> None:
    metadata.create_all(engine)  # cria a tabela se ainda não existir


# ----------------------------------------------------------------------------
# Registros
# ----------------------------------------------------------------------------
def hash_registro(r: dict) -> str:
    """Impressão digital do registro: mesma descrição+valor+fornecedor+vendedor+data+código = mesmo registro."""
    data = r.get("data_cotacao")
    base = "|".join([
        normalizar(r.get("codigo")),
        normalizar(r["descricao"]),
        f"{float(r['valor_unitario']):.4f}",
        normalizar(r.get("fornecedor")),
        normalizar(r.get("vendedor")),
        data.isoformat() if data else "",
    ])
    return hashlib.sha1(base.encode("utf-8")).hexdigest()


def preparar_linhas(registros: list[dict], origem: str, lote: str, agora: datetime) -> dict[str, dict]:
    """Converte os registros extraídos em linhas da tabela, já sem repetições dentro do próprio lote."""
    linhas: dict[str, dict] = {}
    for r in registros:
        h = hash_registro(r)
        if h in linhas:
            continue
        linhas[h] = {
            "codigo": (r.get("codigo") or None) and r["codigo"][:100],
            "descricao": r["descricao"],
            "descricao_norm": normalizar(r["descricao"]),
            "valor_unitario": float(r["valor_unitario"]),
            "fornecedor": (r.get("fornecedor") or None) and r["fornecedor"][:255],
            "vendedor": (r.get("vendedor") or None) and r["vendedor"][:255],
            "data_cotacao": r.get("data_cotacao"),
            "origem": origem[:255],
            "lote": lote,
            "criado_em": agora,
            "hash_linha": h,
        }
    return linhas


def ja_existentes(engine, hashes: list[str]) -> set[str]:
    achados: set[str] = set()
    with engine.connect() as conn:
        for i in range(0, len(hashes), 500):  # em blocos, para não estourar o limite de parâmetros
            bloco = hashes[i:i + 500]
            achados.update(conn.execute(select(precos.c.hash_linha).where(precos.c.hash_linha.in_(bloco))).scalars().all())
    return achados


def classificar(engine, registros: list[dict]) -> list[str]:
    """Para cada registro: Novo, Já no banco ou Repetido na planilha (mesma ordem da lista)."""
    hashes = [hash_registro(r) for r in registros]
    existentes = ja_existentes(engine, list(dict.fromkeys(hashes)))
    vistos: set[str] = set()
    saida = []
    for h in hashes:
        if h in existentes:
            saida.append(SIT_EXISTE)
        elif h in vistos:
            saida.append(SIT_REPETIDO)
        else:
            vistos.add(h)
            saida.append(SIT_NOVO)
    return saida


def inserir(engine, registros: list[dict], origem: str) -> dict:
    """Grava só o que ainda não existe. Retorna {'lote', 'inseridos', 'duplicados'}."""
    lote = uuid.uuid4().hex[:12]
    agora = datetime.now(BRT).replace(tzinfo=None)
    linhas = preparar_linhas(registros, origem, lote, agora)
    existentes = ja_existentes(engine, list(linhas))
    novos = [linha for h, linha in linhas.items() if h not in existentes]
    if novos:
        with engine.begin() as conn:  # transação: ou grava tudo, ou nada
            conn.execute(insert(precos), novos)
    return {"lote": lote, "inseridos": len(novos), "duplicados": len(registros) - len(novos)}


def ler_precos(engine) -> pd.DataFrame:
    with engine.connect() as conn:
        return pd.read_sql(select(precos), conn)


def contar(engine) -> int:
    with engine.connect() as conn:
        return int(conn.execute(select(func.count()).select_from(precos)).scalar_one())


def listar_lotes(engine) -> pd.DataFrame:
    stmt = (
        select(
            precos.c.lote,
            func.max(precos.c.origem).label("origem"),
            func.min(precos.c.criado_em).label("importado_em"),
            func.count().label("registros"),
        )
        .group_by(precos.c.lote)
        .order_by(func.min(precos.c.criado_em).desc())
    )
    with engine.connect() as conn:
        return pd.read_sql(stmt, conn)


def remover_lote(engine, lote: str) -> int:
    with engine.begin() as conn:
        return conn.execute(delete(precos).where(precos.c.lote == lote)).rowcount
