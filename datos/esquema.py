"""Esquema SQLite de la base de datos de cruce de correos."""

VERSION = 1

ESQUEMA = """
CREATE TABLE IF NOT EXISTS carreras (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    nombre_norm TEXT NOT NULL UNIQUE,
    activa INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS periodos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    fecha_inicio TEXT,
    fecha_fin TEXT
);

CREATE TABLE IF NOT EXISTS importaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo TEXT NOT NULL,
    archivo TEXT NOT NULL,
    hoja TEXT,
    hash_archivo TEXT,
    periodo_id INTEGER,
    carrera_id INTEGER,
    fecha TEXT NOT NULL,
    filas INTEGER NOT NULL DEFAULT 0,
    insertados INTEGER NOT NULL DEFAULT 0,
    actualizados INTEGER NOT NULL DEFAULT 0,
    omitidos INTEGER NOT NULL DEFAULT 0,
    errores INTEGER NOT NULL DEFAULT 0,
    estado TEXT NOT NULL DEFAULT 'en_proceso',
    mensaje TEXT,
    FOREIGN KEY(periodo_id) REFERENCES periodos(id),
    FOREIGN KEY(carrera_id) REFERENCES carreras(id)
);

CREATE TABLE IF NOT EXISTS alumnos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    carrera_id INTEGER NOT NULL,
    rut TEXT,
    rut_norm TEXT,
    nombre TEXT NOT NULL,
    nombre_norm TEXT NOT NULL,
    correo TEXT NOT NULL DEFAULT '',
    area TEXT NOT NULL DEFAULT '',
    tipo TEXT NOT NULL DEFAULT '',
    clave TEXT NOT NULL,
    creado_en TEXT NOT NULL,
    actualizado_en TEXT NOT NULL,
    FOREIGN KEY(carrera_id) REFERENCES carreras(id),
    UNIQUE(carrera_id, clave)
);

CREATE TABLE IF NOT EXISTS rotaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    periodo_id INTEGER NOT NULL,
    carrera_id INTEGER NOT NULL,
    codigo TEXT NOT NULL,
    fecha_inicio TEXT,
    fecha_fin TEXT,
    fecha_texto TEXT NOT NULL DEFAULT '',
    fecha_texto_normalizado TEXT NOT NULL DEFAULT '',
    creado_en TEXT NOT NULL,
    actualizado_en TEXT NOT NULL,
    FOREIGN KEY(periodo_id) REFERENCES periodos(id),
    FOREIGN KEY(carrera_id) REFERENCES carreras(id),
    UNIQUE(periodo_id, carrera_id, codigo)
);

CREATE TABLE IF NOT EXISTS campos_clinicos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    carrera_id INTEGER NOT NULL,
    nombre TEXT NOT NULL,
    nombre_norm TEXT NOT NULL,
    nombre_corto TEXT NOT NULL DEFAULT '',
    FOREIGN KEY(carrera_id) REFERENCES carreras(id),
    UNIQUE(carrera_id, nombre_norm)
);

CREATE TABLE IF NOT EXISTS asignaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    rotacion_id INTEGER NOT NULL,
    alumno_id INTEGER NOT NULL,
    campo_id INTEGER NOT NULL,
    orden INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY(rotacion_id) REFERENCES rotaciones(id) ON DELETE CASCADE,
    FOREIGN KEY(alumno_id) REFERENCES alumnos(id),
    FOREIGN KEY(campo_id) REFERENCES campos_clinicos(id),
    UNIQUE(rotacion_id, alumno_id, campo_id)
);

CREATE TABLE IF NOT EXISTS centros_clinicos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    nombre_norm TEXT NOT NULL,
    direccion TEXT NOT NULL DEFAULT '',
    direccion_norm TEXT NOT NULL DEFAULT '',
    contacto TEXT NOT NULL DEFAULT '',
    hora_ingreso TEXT NOT NULL DEFAULT '',
    creado_en TEXT NOT NULL,
    actualizado_en TEXT NOT NULL,
    UNIQUE(nombre_norm, direccion_norm)
);

CREATE TABLE IF NOT EXISTS tutores (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    nombre_norm TEXT NOT NULL,
    correo TEXT NOT NULL DEFAULT '',
    correo_origen TEXT NOT NULL DEFAULT '',
    creado_en TEXT NOT NULL,
    actualizado_en TEXT NOT NULL,
    UNIQUE(nombre_norm, correo)
);

CREATE TABLE IF NOT EXISTS centro_tutores (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    centro_id INTEGER NOT NULL,
    tutor_id INTEGER NOT NULL,
    campo_id INTEGER,
    principal INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY(centro_id) REFERENCES centros_clinicos(id) ON DELETE CASCADE,
    FOREIGN KEY(tutor_id) REFERENCES tutores(id) ON DELETE CASCADE,
    FOREIGN KEY(campo_id) REFERENCES campos_clinicos(id),
    UNIQUE(centro_id, tutor_id, campo_id)
);

CREATE TABLE IF NOT EXISTS catalogo_registros (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    importacion_id INTEGER,
    carrera TEXT NOT NULL DEFAULT '',
    carrera_norm TEXT NOT NULL DEFAULT '',
    fila INTEGER,
    tipo_practica TEXT NOT NULL DEFAULT '',
    area TEXT NOT NULL DEFAULT '',
    campo TEXT NOT NULL DEFAULT '',
    campo_norm TEXT NOT NULL DEFAULT '',
    campo_corto TEXT NOT NULL DEFAULT '',
    campo_corto_norm TEXT NOT NULL DEFAULT '',
    direccion TEXT NOT NULL DEFAULT '',
    contacto TEXT NOT NULL DEFAULT '',
    contacto_nombre TEXT NOT NULL DEFAULT '',
    tutor_texto TEXT NOT NULL DEFAULT '',
    tutor TEXT NOT NULL DEFAULT '',
    tutor_norm TEXT NOT NULL DEFAULT '',
    correo_tutor_texto TEXT NOT NULL DEFAULT '',
    correo_tutor TEXT NOT NULL DEFAULT '',
    correo_tutor_origen TEXT NOT NULL DEFAULT '',
    hora_ingreso TEXT NOT NULL DEFAULT '',
    nombre_residencia TEXT NOT NULL DEFAULT '',
    centro_id INTEGER,
    tutor_id INTEGER,
    campo_id INTEGER,
    FOREIGN KEY(importacion_id) REFERENCES importaciones(id),
    FOREIGN KEY(centro_id) REFERENCES centros_clinicos(id),
    FOREIGN KEY(tutor_id) REFERENCES tutores(id),
    FOREIGN KEY(campo_id) REFERENCES campos_clinicos(id)
);

CREATE TABLE IF NOT EXISTS importacion_filas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    importacion_id INTEGER NOT NULL,
    numero_fila INTEGER,
    datos_json TEXT NOT NULL DEFAULT '{}',
    estado TEXT NOT NULL DEFAULT 'ok',
    mensaje TEXT NOT NULL DEFAULT '',
    FOREIGN KEY(importacion_id) REFERENCES importaciones(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS importacion_incidencias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    importacion_id INTEGER NOT NULL,
    numero_fila INTEGER,
    columna TEXT NOT NULL DEFAULT '',
    valor TEXT NOT NULL DEFAULT '',
    severidad TEXT NOT NULL DEFAULT 'advertencia',
    mensaje TEXT NOT NULL DEFAULT '',
    FOREIGN KEY(importacion_id) REFERENCES importaciones(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_alumnos_nombre
    ON alumnos(carrera_id, nombre_norm);
CREATE INDEX IF NOT EXISTS idx_asignaciones_rotacion
    ON asignaciones(rotacion_id, campo_id);
CREATE INDEX IF NOT EXISTS idx_catalogo_carrera
    ON catalogo_registros(carrera_norm);
CREATE INDEX IF NOT EXISTS idx_importaciones_fecha
    ON importaciones(fecha);

CREATE VIEW IF NOT EXISTS vw_asignaciones AS
SELECT
    p.nombre AS periodo,
    c.nombre AS carrera,
    r.codigo AS rotacion,
    r.fecha_inicio,
    r.fecha_fin,
    r.fecha_texto,
    cc.nombre AS campo,
    a.nombre AS alumno,
    a.rut,
    a.correo,
    a.area,
    a.tipo
FROM asignaciones s
JOIN rotaciones r ON r.id = s.rotacion_id
JOIN periodos p ON p.id = r.periodo_id
JOIN carreras c ON c.id = r.carrera_id
JOIN campos_clinicos cc ON cc.id = s.campo_id
JOIN alumnos a ON a.id = s.alumno_id;

CREATE VIEW IF NOT EXISTS vw_grupos_revision AS
SELECT
    p.nombre AS periodo,
    c.nombre AS carrera,
    r.codigo AS rotacion,
    r.fecha_inicio,
    r.fecha_fin,
    cc.nombre AS campo,
    COUNT(DISTINCT a.id) AS total_alumnos,
    GROUP_CONCAT(a.nombre, ', ') AS alumnos
FROM asignaciones s
JOIN rotaciones r ON r.id = s.rotacion_id
JOIN periodos p ON p.id = r.periodo_id
JOIN carreras c ON c.id = r.carrera_id
JOIN campos_clinicos cc ON cc.id = s.campo_id
JOIN alumnos a ON a.id = s.alumno_id
GROUP BY p.nombre, c.nombre, r.codigo, r.fecha_inicio, r.fecha_fin, cc.nombre;
"""
