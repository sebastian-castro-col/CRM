"""
Gestión de base de datos SQLite para el CRM de empleos.
"""

import sqlite3
import json
import hashlib
import secrets
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "crm_empleos.db"

ESTADOS = [
    "Pendiente",
    "Aplicada",
    "Aplicacion Avanzada Pendiente",
    "Entrevista",
    "Oferta recibida",
    "Descartada",
]
ROLES   = ["admin", "usuario", "visualizador"]

COLOR_ESTADO = {
    "Pendiente":                    "#6c757d",
    "Aplicada":                     "#0d6efd",
    "Aplicacion Avanzada Pendiente":"#6610f2",
    "Entrevista":                   "#fd7e14",
    "Oferta recibida":              "#198754",
    "Descartada":                   "#dc3545",
}
EMOJI_ESTADO = {
    "Pendiente":                    "⏳",
    "Aplicada":                     "📤",
    "Aplicacion Avanzada Pendiente":"🔄",
    "Entrevista":                   "🎤",
    "Oferta recibida":              "🎉",
    "Descartada":                   "🗑️",
}


# ---------------------------------------------------------------------------
# Conexión
# ---------------------------------------------------------------------------

def conectar() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


# ---------------------------------------------------------------------------
# Inicialización
# ---------------------------------------------------------------------------

def _migrar(conn):
    """Agrega columnas nuevas a tablas existentes sin romper datos previos."""
    migraciones = [
        ("busquedas", "ejecutado_por", "TEXT DEFAULT 'sistema'"),
        ("ofertas",   "creado_por",    "TEXT DEFAULT 'sistema'"),
        ("ofertas",   "busqueda_id",   "INTEGER DEFAULT NULL"),
    ]
    cols_cache: dict[str, set] = {}
    for tabla, col, tipo in migraciones:
        if tabla not in cols_cache:
            cols_cache[tabla] = {
                r[1] for r in conn.execute(f"PRAGMA table_info({tabla})").fetchall()
            }
        if col not in cols_cache[tabla]:
            conn.execute(f"ALTER TABLE {tabla} ADD COLUMN {col} {tipo}")
            cols_cache[tabla].add(col)
    conn.commit()


def inicializar():
    with conectar() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS ofertas (
                id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo                  TEXT NOT NULL,
                empresa                 TEXT,
                ubicacion               TEXT,
                modalidad               TEXT,
                descripcion             TEXT,
                url                     TEXT,
                fuente                  TEXT,
                salario                 TEXT DEFAULT 'No especificado',
                compatibilidad_pct      INTEGER DEFAULT 0,
                nivel_compat            TEXT DEFAULT 'Baja',
                tecnologias_encontradas TEXT,
                fecha_publicacion       TEXT,
                fecha_extraccion        TEXT,
                estado                  TEXT DEFAULT 'Pendiente',
                fecha_aplicacion        TEXT,
                notas                   TEXT DEFAULT '',
                fecha_actualizacion     TEXT,
                creado_por              TEXT DEFAULT 'sistema'
            );

            CREATE UNIQUE INDEX IF NOT EXISTS idx_url
                ON ofertas(url) WHERE url IS NOT NULL AND url != '';

            CREATE TABLE IF NOT EXISTS usuarios (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                username         TEXT UNIQUE NOT NULL,
                password_hash    TEXT NOT NULL,
                salt             TEXT NOT NULL,
                nombre_completo  TEXT,
                email            TEXT,
                rol              TEXT DEFAULT 'usuario',
                activo           INTEGER DEFAULT 1,
                fecha_creacion   TEXT,
                ultimo_acceso    TEXT
            );

            CREATE TABLE IF NOT EXISTS busquedas (
                id                INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha             TEXT,
                parametros        TEXT,
                total_nuevas      INTEGER,
                total_encontradas INTEGER,
                ejecutado_por     TEXT
            );
        """)
        conn.commit()
        _migrar(conn)
        # Crear admin por defecto si no existe
        _crear_admin_default(conn)


def _crear_admin_default(conn):
    existe = conn.execute(
        "SELECT COUNT(*) FROM usuarios WHERE username='admin'"
    ).fetchone()[0]
    if not existe:
        salt = secrets.token_hex(16)
        pwd_hash = _hash_password("admin123", salt)
        conn.execute("""
            INSERT INTO usuarios (username, password_hash, salt, nombre_completo, email, rol, fecha_creacion)
            VALUES (?, ?, ?, ?, ?, 'admin', ?)
        """, ("admin", pwd_hash, salt, "Administrador", "admin@crm.local",
              datetime.now().strftime("%Y-%m-%d %H:%M")))
        conn.commit()


# ---------------------------------------------------------------------------
# Autenticación
# ---------------------------------------------------------------------------

def _hash_password(password: str, salt: str) -> str:
    return hashlib.sha256(f"{salt}{password}".encode()).hexdigest()


def autenticar(username: str, password: str) -> dict | None:
    """Retorna dict del usuario si las credenciales son válidas, None si no."""
    with conectar() as conn:
        row = conn.execute(
            "SELECT * FROM usuarios WHERE username=? AND activo=1",
            (username,)
        ).fetchone()
        if not row:
            return None
        if _hash_password(password, row["salt"]) != row["password_hash"]:
            return None
        # Actualizar último acceso
        conn.execute(
            "UPDATE usuarios SET ultimo_acceso=? WHERE id=?",
            (datetime.now().strftime("%Y-%m-%d %H:%M"), row["id"])
        )
        conn.commit()
        return dict(row)


# ---------------------------------------------------------------------------
# Gestión de usuarios (solo admin)
# ---------------------------------------------------------------------------

def crear_usuario(username: str, password: str, nombre: str, email: str, rol: str) -> bool:
    salt = secrets.token_hex(16)
    pwd_hash = _hash_password(password, salt)
    try:
        with conectar() as conn:
            conn.execute("""
                INSERT INTO usuarios (username, password_hash, salt, nombre_completo, email, rol, fecha_creacion)
                VALUES (?,?,?,?,?,?,?)
            """, (username, pwd_hash, salt, nombre, email, rol,
                  datetime.now().strftime("%Y-%m-%d %H:%M")))
            conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False


def obtener_usuarios() -> list[dict]:
    with conectar() as conn:
        rows = conn.execute(
            "SELECT id, username, nombre_completo, email, rol, activo, fecha_creacion, ultimo_acceso "
            "FROM usuarios ORDER BY id"
        ).fetchall()
    return [dict(r) for r in rows]


def actualizar_usuario(user_id: int, nombre: str, email: str, rol: str, activo: bool):
    with conectar() as conn:
        conn.execute("""
            UPDATE usuarios SET nombre_completo=?, email=?, rol=?, activo=?
            WHERE id=?
        """, (nombre, email, rol, int(activo), user_id))
        conn.commit()


def cambiar_password(user_id: int, nueva_password: str):
    salt = secrets.token_hex(16)
    pwd_hash = _hash_password(nueva_password, salt)
    with conectar() as conn:
        conn.execute(
            "UPDATE usuarios SET password_hash=?, salt=? WHERE id=?",
            (pwd_hash, salt, user_id)
        )
        conn.commit()


def eliminar_usuario(user_id: int):
    with conectar() as conn:
        conn.execute("DELETE FROM usuarios WHERE id=? AND username != 'admin'", (user_id,))
        conn.commit()


# ---------------------------------------------------------------------------
# Ofertas
# ---------------------------------------------------------------------------

def insertar_ofertas(
    ofertas: list[dict],
    creado_por: str = "sistema",
    busqueda_id: int = None,
) -> tuple[int, int]:
    """Inserta ofertas. Retorna (insertadas, ignoradas)."""
    insertadas = ignoradas = 0
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    with conectar() as conn:
        for o in ofertas:
            url = (o.get("url") or "").strip() or None
            # Clave de deduplicación: titulo[:60] + empresa[:40] + fuente
            # Previene duplicados cuando la URL cambia (tracking params, etc.)
            titulo_key  = (o.get("titulo", "") or "")[:60].lower()
            empresa_key = (o.get("empresa", "") or "")[:40].lower()
            fuente_key  = o.get("fuente", "")
            ya_existe = conn.execute("""
                SELECT COUNT(*) FROM ofertas
                WHERE LOWER(SUBSTR(titulo,  1, 60)) = ?
                  AND LOWER(SUBSTR(empresa, 1, 40)) = ?
                  AND fuente = ?
            """, (titulo_key, empresa_key, fuente_key)).fetchone()[0]
            if ya_existe:
                ignoradas += 1
                continue
            try:
                result = conn.execute("""
                    INSERT OR IGNORE INTO ofertas
                    (titulo, empresa, ubicacion, modalidad, descripcion, url, fuente,
                     salario, compatibilidad_pct, nivel_compat, tecnologias_encontradas,
                     fecha_publicacion, fecha_extraccion, creado_por, busqueda_id)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    o.get("titulo", ""), o.get("empresa", ""), o.get("ubicacion", ""),
                    o.get("modalidad", "Remoto"), o.get("descripcion", ""),
                    url, o.get("fuente", ""),
                    o.get("salario", "No especificado"),
                    o.get("compatibilidad_pct", 0), o.get("nivel_compat", "Baja"),
                    o.get("tecnologias_encontradas", ""),
                    o.get("fecha_publicacion", ""),
                    o.get("fecha_extraccion", now),
                    creado_por, busqueda_id,
                ))
                if result.rowcount:
                    insertadas += 1
                else:
                    ignoradas += 1
            except Exception:
                ignoradas += 1
        conn.commit()
    return insertadas, ignoradas


def obtener_ofertas(
    estado: list[str] = None,
    fuentes: list[str] = None,
    nivel_compat: list[str] = None,
    modalidad: list[str] = None,
    busqueda_texto: str = None,
    fecha_desde: str = None,
    fecha_hasta: str = None,
    orden: str = "compatibilidad_pct DESC",
    usuario: str = None,          # Aislamiento por usuario
) -> list[dict]:
    where, params = [], []

    # Deduplicación automática: un registro por (titulo[:60], empresa[:40], fuente).
    # Prefiere el que NO está Descartada (MAX id no-Descartada); si todos lo están, MIN id.
    # La subquery usa el mismo filtro de usuario para evitar cruce entre cuentas.
    if usuario:
        where.append("""
            id IN (
                SELECT COALESCE(
                    MAX(CASE WHEN estado != 'Descartada' THEN id END),
                    MIN(id)
                )
                FROM ofertas
                WHERE creado_por = ?
                GROUP BY LOWER(SUBSTR(titulo,  1, 60)),
                         LOWER(SUBSTR(empresa, 1, 40)),
                         fuente
            )
        """)
        params.append(usuario)
        # Filtro exterior también (seguridad extra)
        where.append("creado_por = ?")
        params.append(usuario)
    else:
        where.append("""
            id IN (
                SELECT COALESCE(
                    MAX(CASE WHEN estado != 'Descartada' THEN id END),
                    MIN(id)
                )
                FROM ofertas
                GROUP BY LOWER(SUBSTR(titulo,  1, 60)),
                         LOWER(SUBSTR(empresa, 1, 40)),
                         fuente
            )
        """)

    if estado:
        where.append(f"estado IN ({','.join('?'*len(estado))})")
        params.extend(estado)
    if fuentes:
        where.append(f"fuente IN ({','.join('?'*len(fuentes))})")
        params.extend(fuentes)
    if nivel_compat:
        where.append(f"nivel_compat IN ({','.join('?'*len(nivel_compat))})")
        params.extend(nivel_compat)
    if modalidad:
        where.append(f"modalidad IN ({','.join('?'*len(modalidad))})")
        params.extend(modalidad)
    if busqueda_texto:
        where.append("(titulo LIKE ? OR empresa LIKE ? OR tecnologias_encontradas LIKE ?)")
        t = f"%{busqueda_texto}%"
        params.extend([t, t, t])
    if fecha_desde:
        where.append("fecha_extraccion >= ?")
        params.append(str(fecha_desde))
    if fecha_hasta:
        where.append("fecha_extraccion <= ?")
        params.append(str(fecha_hasta) + " 23:59")

    sql = "SELECT * FROM ofertas"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += f" ORDER BY {orden}"

    with conectar() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [dict(r) for r in rows]


def actualizar_estado(oferta_id: int, estado: str, notas: str = None):
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    with conectar() as conn:
        fecha_ap = now if estado == "Aplicada" else None
        if fecha_ap:
            conn.execute(
                "UPDATE ofertas SET estado=?, fecha_aplicacion=?, fecha_actualizacion=? WHERE id=?",
                (estado, fecha_ap, now, oferta_id)
            )
        else:
            conn.execute(
                "UPDATE ofertas SET estado=?, fecha_actualizacion=? WHERE id=?",
                (estado, now, oferta_id)
            )
        if notas is not None:
            conn.execute("UPDATE ofertas SET notas=? WHERE id=?", (notas, oferta_id))
        conn.commit()


def estadisticas(usuario: str = None) -> dict:
    filtro = "WHERE creado_por=?" if usuario else ""
    p = (usuario,) if usuario else ()
    with conectar() as conn:
        total    = conn.execute(f"SELECT COUNT(*) FROM ofertas {filtro}", p).fetchone()[0]
        p_estado = dict(conn.execute(f"SELECT estado, COUNT(*) FROM ofertas {filtro} GROUP BY estado", p).fetchall())
        p_fuente = dict(conn.execute(f"SELECT fuente, COUNT(*) FROM ofertas {filtro} GROUP BY fuente", p).fetchall())
        p_nivel  = dict(conn.execute(f"SELECT nivel_compat, COUNT(*) FROM ofertas {filtro} GROUP BY nivel_compat", p).fetchall())
        p_modal  = dict(conn.execute(f"SELECT modalidad, COUNT(*) FROM ofertas {filtro} GROUP BY modalidad", p).fetchall())
        prom     = conn.execute(f"SELECT ROUND(AVG(compatibilidad_pct),1) FROM ofertas {filtro}", p).fetchone()[0] or 0
    return {
        "total": total, "por_estado": p_estado, "por_fuente": p_fuente,
        "por_nivel": p_nivel, "por_modalidad": p_modal, "prom_compat": prom,
    }


def registrar_busqueda(
    parametros: dict,
    total_nuevas: int,
    total_encontradas: int,
    usuario: str = "sistema",
) -> int:
    """Registra una búsqueda y retorna su ID."""
    with conectar() as conn:
        cur = conn.execute("""
            INSERT INTO busquedas (fecha, parametros, total_nuevas, total_encontradas, ejecutado_por)
            VALUES (?,?,?,?,?)
        """, (
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            json.dumps(parametros, ensure_ascii=False),
            total_nuevas, total_encontradas, usuario,
        ))
        conn.commit()
        return cur.lastrowid


def historial_busquedas(limit: int = 15, usuario: str = None) -> list[dict]:
    with conectar() as conn:
        if usuario:
            rows = conn.execute(
                "SELECT * FROM busquedas WHERE ejecutado_por=? ORDER BY fecha DESC LIMIT ?",
                (usuario, limit)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM busquedas ORDER BY fecha DESC LIMIT ?", (limit,)
            ).fetchall()
    return [dict(r) for r in rows]


def ofertas_por_busqueda(busqueda_id: int, usuario: str = None) -> list[dict]:
    """Retorna ofertas de una búsqueda, opcionalmente filtradas por usuario."""
    with conectar() as conn:
        if usuario:
            rows = conn.execute(
                "SELECT * FROM ofertas WHERE busqueda_id=? AND creado_por=? "
                "ORDER BY compatibilidad_pct DESC",
                (busqueda_id, usuario)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM ofertas WHERE busqueda_id=? ORDER BY compatibilidad_pct DESC",
                (busqueda_id,)
            ).fetchall()
    return [dict(r) for r in rows]


def buscar_por_job_id(job_id: int, usuario: str = None) -> dict | None:
    """Retorna una oferta por ID, verificando que pertenezca al usuario."""
    with conectar() as conn:
        if usuario:
            row = conn.execute(
                "SELECT * FROM ofertas WHERE id=? AND creado_por=?", (job_id, usuario)
            ).fetchone()
        else:
            row = conn.execute("SELECT * FROM ofertas WHERE id=?", (job_id,)).fetchone()
    return dict(row) if row else None


def descartar_ofertas_ingles(palabras_ingles: list[str]) -> int:
    """
    Busca ofertas cuya descripción/título contenga palabras de inglés
    y las marca como 'Descartada'. Retorna cuántas se descartaron.
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    descartadas = 0
    with conectar() as conn:
        rows = conn.execute(
            "SELECT id, titulo, descripcion FROM ofertas WHERE estado != 'Descartada'"
        ).fetchall()
        for row in rows:
            texto = ((row["titulo"] or "") + " " + (row["descripcion"] or "")).lower()
            # Chequear excepciones primero
            excepciones = [
                "no se requiere inglés", "no requiere inglés",
                "inglés no requerido", "no es necesario inglés",
                "inglés no es", "no exige inglés",
            ]
            if any(exc in texto for exc in excepciones):
                continue
            for pal in palabras_ingles:
                if pal in texto:
                    conn.execute(
                        "UPDATE ofertas SET estado='Descartada', fecha_actualizacion=? WHERE id=?",
                        (now, row["id"])
                    )
                    descartadas += 1
                    break
        conn.commit()
    return descartadas


def limpiar_duplicados() -> int:
    """
    Elimina ofertas duplicadas por (titulo[:60], empresa[:40], fuente).
    Conserva el registro que NO esté Descartada (mayor id); si todos están
    Descartados conserva el más antiguo. Retorna cuántos se eliminaron.
    """
    with conectar() as conn:
        cur = conn.execute("""
            DELETE FROM ofertas
            WHERE id NOT IN (
                SELECT COALESCE(
                    MAX(CASE WHEN estado != 'Descartada' THEN id END),
                    MIN(id)
                )
                FROM   ofertas
                GROUP BY
                    LOWER(SUBSTR(titulo,  1, 60)),
                    LOWER(SUBSTR(empresa, 1, 40)),
                    fuente
            )
        """)
        conn.commit()
        return cur.rowcount


def descartar_presenciales(palabras_presencial: list[str]) -> int:
    """
    Busca ofertas presenciales en título/descripción y las marca como 'Descartada'.
    Retorna cuántas se descartaron.
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    descartadas = 0
    with conectar() as conn:
        rows = conn.execute(
            "SELECT id, titulo, descripcion FROM ofertas WHERE estado != 'Descartada'"
        ).fetchall()
        for row in rows:
            texto = ((row["titulo"] or "") + " " + (row["descripcion"] or "")).lower()
            for pal in palabras_presencial:
                if pal in texto:
                    conn.execute(
                        "UPDATE ofertas SET estado='Descartada', notas=?, fecha_actualizacion=? WHERE id=?",
                        (f"Auto-descartada: modalidad presencial ({pal})", now, row["id"])
                    )
                    descartadas += 1
                    break
        conn.commit()
    return descartadas


def descartar_empresa(empresa: str) -> int:
    """Descarta todas las ofertas de una empresa específica."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    with conectar() as conn:
        cur = conn.execute(
            "UPDATE ofertas SET estado='Descartada', notas=?, fecha_actualizacion=? "
            "WHERE LOWER(empresa) LIKE ? AND estado != 'Descartada'",
            (f"Auto-descartada: empresa excluida", now, f"%{empresa.lower()}%")
        )
        conn.commit()
        return cur.rowcount


def eliminar_todas_ofertas() -> int:
    """Elimina todas las ofertas de la BD. Retorna cantidad eliminada."""
    with conectar() as conn:
        n = conn.execute("SELECT COUNT(*) FROM ofertas").fetchone()[0]
        conn.execute("DELETE FROM ofertas")
        conn.commit()
    return n


inicializar()
