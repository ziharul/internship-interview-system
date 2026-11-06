"""PostgreSQL persistence for Internship Hub.

Database requests run only on the Streamlit server. Never expose DATABASE_URL
in GitHub, browser code, logs, or downloadable artifacts.
"""
from __future__ import annotations
from datetime import date, datetime
from decimal import Decimal
from typing import Any

import psycopg
from psycopg import sql
from psycopg.rows import dict_row
import streamlit as st

from data_store import KANDIDAT_FIELDS, ASPEK

CANDIDATE_KEYS = [key for key, _ in KANDIDAT_FIELDS]
INTERVIEW_KEYS = [
    'id', 'tanggal', 'batch', *[key for key, _, _ in ASPEK],
    'keputusan', 'kelebihan', 'peningkatan', 'catatan', 'pewawancara',
]
CANDIDATE_NUMBERS = {'semester', 'durasi', 'ipk'}
CANDIDATE_DATES = {'mulai'}
INTERVIEW_DATES = {'tanggal'}
INTERVIEW_SCORES = {key for key, _, _ in ASPEK}


class DatabaseOperationError(Exception):
    """Safe database error message for application UI."""


def configured() -> bool:
    try:
        return bool(st.secrets.get('DATABASE_URL'))
    except Exception:
        return False


def _connect():
    if not configured():
        raise DatabaseOperationError('DATABASE_URL belum diatur di Settings → Secrets.')
    try:
        # Fresh, short-lived connection for each operation: no stale session across reruns.
        return psycopg.connect(str(st.secrets['DATABASE_URL']), connect_timeout=12, sslmode='require', row_factory=dict_row)
    except psycopg.Error as exc:
        raise DatabaseOperationError('Tidak dapat tersambung ke PostgreSQL. Periksa Secrets dan mode Session pooler.') from exc


def _to_text(value: Any) -> str:
    if value is None:
        return ''
    if isinstance(value, (date, datetime)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, Decimal):
        return format(value, 'f')
    return str(value)


def _to_candidate_value(key: str, value: Any):
    if value is None or value == '':
        return None
    if key in ('semester', 'durasi'):
        return int(value)
    if key == 'ipk':
        return Decimal(str(value))
    if key == 'mulai':
        return value if isinstance(value, date) else date.fromisoformat(str(value)[:10])
    return str(value)


def _to_interview_value(key: str, value: Any):
    if key in INTERVIEW_SCORES:
        return int(value) if value not in (None, '', 0, '0') else None
    if key == 'tanggal':
        return date.fromisoformat(str(value)[:10]) if value else None
    if key == 'batch':
        if value in (None, ''):
            return None  # Preserve unassigned historical records.
        from data_store import BATCH_PATTERN
        if not BATCH_PATTERN.fullmatch(str(value)):
            raise ValueError('Format batch tidak valid. Gunakan format YYYY-BNN.')
        return str(value)
    if key == 'keputusan':
        return str(value or 'Menunggu')
    return str(value) if value not in (None, '') else None


def load_records() -> tuple[dict[str, dict], dict[str, dict]]:
    try:
        with _connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql.SQL('SELECT {} FROM magang.kandidat ORDER BY id').format(
                    sql.SQL(', ').join(map(sql.Identifier, CANDIDATE_KEYS))))
                candidates = {}
                for row in cur.fetchall():
                    rec = {key: _to_text(row.get(key)) for key in CANDIDATE_KEYS}
                    candidates[rec['id']] = rec
                cur.execute(sql.SQL('SELECT {} FROM magang.wawancara ORDER BY id').format(
                    sql.SQL(', ').join(map(sql.Identifier, INTERVIEW_KEYS))))
                interviews = {}
                for row in cur.fetchall():
                    rec = {
                        key: (int(row[key]) if row.get(key) is not None else 0)
                        if key in INTERVIEW_SCORES else _to_text(row.get(key))
                        for key in INTERVIEW_KEYS
                    }
                    interviews[rec['id']] = rec
                return candidates, interviews
    except DatabaseOperationError:
        raise
    except psycopg.Error as exc:
        raise DatabaseOperationError('Gagal membaca tabel magang. Pastikan setup_database.sql sudah dijalankan.') from exc


def _write_candidate(cur, record: dict, *, create: bool = False, upsert: bool = False):
    vals = [_to_candidate_value(key, record.get(key)) for key in CANDIDATE_KEYS]
    idents = sql.SQL(', ').join(map(sql.Identifier, CANDIDATE_KEYS))
    placeholders = sql.SQL(', ').join([sql.Placeholder()] * len(CANDIDATE_KEYS))
    if create:
        stmt = sql.SQL('INSERT INTO magang.kandidat ({}) VALUES ({})').format(idents, placeholders)
        cur.execute(stmt, vals)
    elif upsert:
        assignments = sql.SQL(', ').join([
            sql.SQL('{} = EXCLUDED.{}').format(sql.Identifier(key), sql.Identifier(key))
            for key in CANDIDATE_KEYS if key != 'id'
        ])
        stmt = sql.SQL('INSERT INTO magang.kandidat ({}) VALUES ({}) '
                       'ON CONFLICT (id) DO UPDATE SET {}, updated_at = now()').format(
            idents, placeholders, assignments)
        cur.execute(stmt, vals)
    else:
        assignments = sql.SQL(', ').join([
            sql.SQL('{} = {}').format(sql.Identifier(key), sql.Placeholder())
            for key in CANDIDATE_KEYS if key != 'id'
        ])
        stmt = sql.SQL('UPDATE magang.kandidat SET {}, updated_at = now() WHERE id = {}').format(
            assignments, sql.Placeholder())
        cur.execute(stmt, [*_to_candidate_values_without_id(record), record['id']])
        if cur.rowcount != 1:
            raise DatabaseOperationError('Kandidat tidak ditemukan. Muat ulang halaman dan coba lagi.')


def _to_candidate_values_without_id(record: dict) -> list[Any]:
    return [_to_candidate_value(key, record.get(key)) for key in CANDIDATE_KEYS if key != 'id']


def save_candidate(record: dict, *, create: bool):
    try:
        with _connect() as conn:
            with conn.cursor() as cur:
                _write_candidate(cur, record, create=create)
    except psycopg.errors.UniqueViolation as exc:
        raise DatabaseOperationError('ID kandidat sudah digunakan. Muat ulang halaman dan coba tambah kandidat kembali.') from exc
    except (psycopg.Error, ValueError, TypeError) as exc:
        raise DatabaseOperationError('Data kandidat gagal disimpan. Periksa IPK, semester, tanggal, dan koneksi database.') from exc


def _write_interview(cur, record: dict, *, preserve_existing_batch: bool = False):
    fields = INTERVIEW_KEYS
    vals = [_to_interview_value(key, record.get(key)) for key in fields]
    identifiers = sql.SQL(', ').join(map(sql.Identifier, fields))
    placeholders = sql.SQL(', ').join([sql.Placeholder()] * len(fields))
    assignments = sql.SQL(', ').join([
        # An older Excel backup has no batch column. Importing it must not erase
        # a batch already recorded in PostgreSQL for the same candidate.
        (sql.SQL('batch = COALESCE(EXCLUDED.batch, magang.wawancara.batch)')
         if k == 'batch' and preserve_existing_batch else
         sql.SQL('{} = EXCLUDED.{}').format(sql.Identifier(k), sql.Identifier(k)))
        for k in fields if k != 'id'
    ])
    cur.execute(sql.SQL('INSERT INTO magang.wawancara ({}) VALUES ({}) '
                        'ON CONFLICT (id) DO UPDATE SET {}, updated_at = now()').format(
        identifiers, placeholders, assignments), vals)


def save_interview(record: dict):
    try:
        with _connect() as conn:
            with conn.cursor() as cur:
                _write_interview(cur, record)
    except (psycopg.Error, ValueError, TypeError) as exc:
        raise DatabaseOperationError('Hasil interview gagal disimpan. Periksa nilai dan koneksi database.') from exc


def import_backup(candidates: dict, interviews: dict):
    """Atomic merge/upsert. Never silently delete existing database data."""
    if any(cid not in candidates for cid in interviews):
        raise DatabaseOperationError('File backup tidak valid: wawancara tanpa kandidat.')
    try:
        with _connect() as conn:
            with conn.cursor() as cur:
                for rec in candidates.values():
                    _write_candidate(cur, rec, upsert=True)
                for rec in interviews.values():
                    _write_interview(cur, rec, preserve_existing_batch=True)
    except (psycopg.Error, ValueError, TypeError) as exc:
        raise DatabaseOperationError('Impor database gagal, tidak ada perubahan parsial yang disimpan.') from exc
