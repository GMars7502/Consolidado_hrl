"""
Logica del Sistema Consolidador (sin interfaz grafica).

Basado en procesar_todo.py: mismas reglas de homologacion y mismo formato final,
pero con gestion propia de carpetas (MAESTROS, TRAMAS, SALIDA) junto al ejecutable,
registro de fechas de actualizacion y soporte de CSV, ZIP y Excel.
"""
import glob
import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from datetime import datetime

import pandas as pd

# ----------------------------------------------------------------------------
# Rutas
# ----------------------------------------------------------------------------

def base_dir():
    """Carpeta del .exe (o del script si se ejecuta con Python)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class Rutas:
    def __init__(self, base=None):
        self.base = base or base_dir()
        self.maestros = os.path.join(self.base, "MAESTROS")
        self.tablas = os.path.join(self.maestros, "TABLAS")
        self.tramas = os.path.join(self.base, "TRAMAS")
        self.salida = os.path.join(self.base, "SALIDA")
        self.registro = os.path.join(self.maestros, "registro_maestros.json")
        self.excel_legado = os.path.join(self.maestros, "TablasMaestrasV2.xlsx")

    def crear(self):
        for d in (self.maestros, self.tablas, self.tramas, self.salida):
            os.makedirs(d, exist_ok=True)
        for sub in ("PACIENTES", "PERSONAL", "REGISTRADORES"):
            os.makedirs(os.path.join(self.maestros, sub), exist_ok=True)


# ----------------------------------------------------------------------------
# Definicion de maestros
# ----------------------------------------------------------------------------

COLS_PACIENTE = ['Id_Paciente', 'Id_Tipo_Documento', 'Numero_Documento', 'Apellido_Paterno_Paciente',
                 'Apellido_Materno_Paciente', 'Nombres_Paciente', 'Fecha_Nacimiento', 'Genero', 'Id_Etnia',
                 'Historia_Clinica']
COLS_PERSONAL = ['Id_Personal', 'Id_Tipo_Documento', 'Numero_Documento', 'Apellido_Paterno_Personal',
                 'Apellido_Materno_Personal', 'Nombres_Personal', 'Fecha_Nacimiento', 'Id_Condicion',
                 'Id_Profesion']
COLS_REGISTRADOR = ['Id_Registrador', 'Id_Tipo_Documento', 'Numero_Documento', 'Apellido_Paterno_Registrador',
                    'Apellido_Materno_Registrador', 'Nombres_Registrador']

# tipo: carpeta (un CSV dentro de una subcarpeta), archivo (CSV fijo), tabla (hoja del Excel guardada como CSV)
MAESTROS = {
    "pacientes": dict(nombre="Pacientes", tipo="carpeta", carpeta="PACIENTES", cols=COLS_PACIENTE, zip=True),
    "registradores": dict(nombre="Registradores", tipo="carpeta", carpeta="REGISTRADORES",
                          cols=COLS_REGISTRADOR, zip=True),
    "personal": dict(nombre="Personal", tipo="carpeta", carpeta="PERSONAL", cols=COLS_PERSONAL, zip=True),
    "cie10": dict(nombre="Catalogo CIE10", tipo="archivo", archivo="catalogo_cie10.csv",
                  cols=["Codigo_Item", "Descripcion_Item"], zip=False),
    "cpms": dict(nombre="Catalogo CPMS", tipo="archivo", archivo="catalogo_cpms.csv",
                 cols=["Codigo_Item", "Descripcion_Item"], zip=False),
    "ups": dict(nombre="Tabla UPS", tipo="tabla", patrones=["UPS"], archivo="ups.csv",
                id=["id_ups", "cod_ups", "ups"], desc=["desc_ups", "descripcion_ups", "ups_desc"], zip=False),
    "establecimiento": dict(nombre="Tabla Establecimiento", tipo="tabla", patrones=["ESTABLECIMIENTO"],
                            archivo="establecimiento.csv",
                            id=["codRenipress", "id_establecimiento", "cod_establecimiento"],
                            desc=["Establecimiento", "nombre_establecimiento", "desc_establecimiento"],
                            zip=False),
    "condicion": dict(nombre="Tabla Condicion", tipo="tabla", patrones=["CONDICION"], archivo="condicion.csv",
                      id=["id_condicion", "cod_condicion"],
                      desc=["desc_condicion", "condicion", "descripcion_condicion"], zip=False),
    "profesion": dict(nombre="Tabla Profesion", tipo="tabla", patrones=["PROFESION"], archivo="profesion.csv",
                      id=["id_profesion", "cod_profesion"],
                      desc=["Profesion", "desc_profesion", "descripcion_profesion"], zip=False),
}
CLAVES_TABLAS = [k for k, v in MAESTROS.items() if v["tipo"] == "tabla"]

COLUMNAS_FINALES = [
    'Id_Cita', 'Anio', 'Mes', 'Dia', 'Fecha_Atencion', 'Lote', 'Num_Pag', 'Num_Reg',
    'Ups', 'Establecimiento', 'Id_Paciente', 'Id_Personal', 'Id_Registrador',
    'Id_Financiador', 'Id_Condicion_Establecimiento', 'Id_Condicion_Servicio',
    'Edad_Reg', 'Tipo_Edad', 'Anio_Actual_Paciente', 'Mes_Actual_Paciente',
    'Dia_Actual_Paciente', 'Id_Turno', 'Codigo_Item', 'Tipo_Diagnostico',
    'Valor_Lab', 'Id_Correlativo', 'Id_Correlativo_Lab', 'Id_Otra_Condicion',
    'Id_Centro_Poblado', 'Fecha_Ultima_Regla', 'Fecha_Registro', 'Fecha_Modificacion',
    'Id_dosis', 'Descripcion_Item',
    'Id_Tipo_Documento_Paciente', 'Numero_Documento_Paciente', 'Apellido_Paterno_Paciente',
    'Apellido_Materno_Paciente', 'Nombres_Paciente', 'Fecha_Nacimiento_Paciente', 'Genero', 'Id_Etnia',
    'Historia_Clinica', 'Id_Pais',
    'Id_Tipo_Documento_Personal', 'Numero_Documento_Personal', 'Apellido_Paterno_Personal',
    'Apellido_Materno_Personal', 'Nombres_Personal', 'Fecha_Nacimiento_Personal', 'Condicion', 'Profesion',
    'Id_Tipo_Documento_Registrador', 'Numero_Documento_Registrador', 'Apellido_Paterno_Registrador',
    'Apellido_Materno_Registrador', 'Nombres_Registrador'
]


class ErrorConsolidador(Exception):
    pass


# ----------------------------------------------------------------------------
# Utilidades (misma logica de procesar_todo.py)
# ----------------------------------------------------------------------------

def clean_key(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    try:
        s = str(int(s))
    except ValueError:
        pass
    return s


def agregar_a_mapa(dicc, key, val):
    if pd.isna(key) or not str(key).strip():
        return
    k_raw = str(key).strip()
    if k_raw.endswith(".0"):
        k_raw = k_raw[:-2]
    v = str(val).strip() if pd.notna(val) else ""
    dicc[k_raw] = v
    try:
        k_int = str(int(k_raw))
        dicc[k_int] = v
    except ValueError:
        pass


def buscar_hoja(sheet_names, patrones):
    for s in sheet_names:
        s_upper = s.upper()
        if any(p.upper() in s_upper for p in patrones):
            return s
    raise ErrorConsolidador(f"No se encontro ninguna hoja que coincida con {patrones} entre {sheet_names}")


def obtener_columna(columnas, candidatos):
    cols_lower = {str(c).strip().lower(): c for c in columnas}
    for cand in candidatos:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    raise ErrorConsolidador(f"No se encontro ninguna columna coincidente con {candidatos} en {list(columnas)}")


def read_csv_safe(ruta, **kw):
    try:
        return pd.read_csv(ruta, dtype=str, encoding="utf-8-sig", **kw)
    except UnicodeDecodeError:
        return pd.read_csv(ruta, dtype=str, encoding="latin-1", **kw)


def buscar_csv_en_carpeta(carpeta):
    archivos = glob.glob(os.path.join(carpeta, "**", "*.csv"), recursive=True)
    return archivos[0] if archivos else None


def contar_filas(ruta):
    """Cuenta registros (lineas - encabezado) leyendo por bloques."""
    n = 0
    ultimo = b""
    with open(ruta, "rb") as f:
        while True:
            bloque = f.read(1 << 20)
            if not bloque:
                break
            n += bloque.count(b"\n")
            ultimo = bloque[-1:]
    if ultimo and ultimo != b"\n":
        n += 1
    return max(n - 1, 0)


def nombre_seguro(nombre):
    nombre = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nombre.strip()).strip(". ")
    if not nombre:
        raise ErrorConsolidador("El nombre del consolidado no puede estar vacio.")
    if not nombre.lower().endswith(".csv"):
        nombre += ".csv"
    return nombre


def extraer_csvs_zip(ruta_zip, destino):
    """Extrae todos los CSV de un ZIP (aplanados) a 'destino'. Devuelve las rutas extraidas."""
    extraidos = []
    try:
        with zipfile.ZipFile(ruta_zip) as z:
            for info in z.infolist():
                if info.is_dir() or not info.filename.lower().endswith(".csv"):
                    continue
                if "__MACOSX" in info.filename:
                    continue
                dest = os.path.join(destino, os.path.basename(info.filename))
                with z.open(info) as src, open(dest, "wb") as out:
                    shutil.copyfileobj(src, out)
                extraidos.append(dest)
    except zipfile.BadZipFile:
        raise ErrorConsolidador(f"'{os.path.basename(ruta_zip)}' no es un ZIP valido.")
    if not extraidos:
        raise ErrorConsolidador(f"El ZIP '{os.path.basename(ruta_zip)}' no contiene ningun archivo CSV.")
    return extraidos


# ----------------------------------------------------------------------------
# Gestor de maestros
# ----------------------------------------------------------------------------

class Gestor:
    def __init__(self, rutas=None):
        self.rutas = rutas or Rutas()
        self.rutas.crear()

    # --- registro de fechas ---
    def _leer_registro(self):
        try:
            with open(self.rutas.registro, "r", encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}

    def _registrar(self, clave, origen, registros, fecha=None):
        reg = self._leer_registro()
        reg[clave] = {
            "fecha": (fecha or datetime.now()).isoformat(timespec="seconds"),
            "origen": origen,
            "registros": registros,
        }
        with open(self.rutas.registro, "w", encoding="utf-8") as f:
            json.dump(reg, f, ensure_ascii=False, indent=2)

    # --- rutas de archivos ---
    def ruta_maestro(self, clave):
        spec = MAESTROS[clave]
        if spec["tipo"] == "carpeta":
            return buscar_csv_en_carpeta(os.path.join(self.rutas.maestros, spec["carpeta"]))
        if spec["tipo"] == "archivo":
            r = os.path.join(self.rutas.maestros, spec["archivo"])
        else:
            r = os.path.join(self.rutas.tablas, spec["archivo"])
        return r if os.path.exists(r) else None

    def estado(self):
        """Lista de dicts con el estado de cada maestro."""
        reg = self._leer_registro()
        out = []
        for clave, spec in MAESTROS.items():
            ruta = self.ruta_maestro(clave)
            info = reg.get(clave) if ruta else None
            out.append(dict(
                clave=clave, nombre=spec["nombre"], existe=bool(ruta),
                fecha=info["fecha"] if info else None,
                origen=info.get("origen") if info else None,
                registros=info.get("registros") if info else None,
            ))
        return out

    def sincronizar(self, log=print):
        """Registra maestros ya existentes sin registro e importa el Excel legado si hace falta."""
        reg = self._leer_registro()
        # Excel legado (TablasMaestrasV2.xlsx) -> CSV de tablas
        faltan = [k for k in CLAVES_TABLAS if not self.ruta_maestro(k)]
        if faltan and os.path.exists(self.rutas.excel_legado):
            log(f"Importando tablas desde {os.path.basename(self.rutas.excel_legado)} (solo la primera vez)...")
            try:
                self.importar_excel_tablas(self.rutas.excel_legado, log=log, claves=faltan)
            except ErrorConsolidador as e:
                log(f"ADVERTENCIA: {e}")
            reg = self._leer_registro()
        for clave in MAESTROS:
            ruta = self.ruta_maestro(clave)
            if ruta and clave not in reg:
                self._registrar(clave, os.path.basename(ruta), contar_filas(ruta),
                                fecha=datetime.fromtimestamp(os.path.getmtime(ruta)))

    # --- actualizacion ---
    def _validar(self, clave, columnas):
        spec = MAESTROS[clave]
        if spec["tipo"] == "tabla":
            obtener_columna(columnas, spec["id"])
            obtener_columna(columnas, spec["desc"])
            return
        faltantes = [c for c in spec["cols"] if c not in list(columnas)]
        if faltantes:
            raise ErrorConsolidador(
                f"Al archivo le faltan columnas requeridas para {spec['nombre']}: {', '.join(faltantes)}")

    def _excel_a_csv(self, clave, ruta_excel, destino_csv):
        spec = MAESTROS[clave]
        xls = pd.ExcelFile(ruta_excel)
        if spec["tipo"] == "tabla":
            hoja = buscar_hoja(xls.sheet_names, spec["patrones"])
        else:
            hoja = xls.sheet_names[0]
        df = pd.read_excel(xls, hoja, dtype=str)
        df.to_csv(destino_csv, index=False, encoding="utf-8")
        return hoja

    def actualizar(self, clave, ruta_origen, log=print):
        """Actualiza un maestro desde CSV, ZIP o Excel."""
        spec = MAESTROS[clave]
        ext = os.path.splitext(ruta_origen)[1].lower()
        base = os.path.basename(ruta_origen)
        with tempfile.TemporaryDirectory() as tmp:
            if ext == ".zip":
                if not spec["zip"]:
                    raise ErrorConsolidador(f"{spec['nombre']} no admite archivos ZIP.")
                log(f"Extrayendo {base}...")
                csvs = extraer_csvs_zip(ruta_origen, tmp)
                src = csvs[0]
                nombre_csv = os.path.basename(src)
            elif ext in (".xlsx", ".xlsm", ".xls"):
                log(f"Leyendo Excel {base} (puede tardar)...")
                nombre_csv = os.path.splitext(base)[0] + ".csv"
                src = os.path.join(tmp, nombre_csv)
                self._excel_a_csv(clave, ruta_origen, src)
            elif ext in (".csv", ".txt"):
                src = ruta_origen
                nombre_csv = base
            else:
                raise ErrorConsolidador(f"Formato no soportado: {ext}")

            log("Validando columnas...")
            self._validar(clave, read_csv_safe(src, nrows=0).columns)

            if spec["tipo"] == "carpeta":
                carpeta = os.path.join(self.rutas.maestros, spec["carpeta"])
                shutil.rmtree(carpeta, ignore_errors=True)
                os.makedirs(carpeta, exist_ok=True)
                dest = os.path.join(carpeta, nombre_csv)
            elif spec["tipo"] == "archivo":
                dest = os.path.join(self.rutas.maestros, spec["archivo"])
            else:
                dest = os.path.join(self.rutas.tablas, spec["archivo"])
            shutil.copyfile(src, dest)

        n = contar_filas(dest)
        self._registrar(clave, base, n)
        log(f"{spec['nombre']} actualizado: {n:,} registros.")

    def importar_excel_tablas(self, ruta_excel, log=print, claves=None):
        """Actualiza UPS, Establecimiento, Condicion y Profesion desde un Excel completo."""
        claves = claves or CLAVES_TABLAS
        xls = pd.ExcelFile(ruta_excel)
        base = os.path.basename(ruta_excel)
        hechos = 0
        for clave in claves:
            spec = MAESTROS[clave]
            try:
                hoja = buscar_hoja(xls.sheet_names, spec["patrones"])
            except ErrorConsolidador:
                log(f"ADVERTENCIA: el Excel no tiene hoja para {spec['nombre']}; se omite.")
                continue
            df = pd.read_excel(xls, hoja, dtype=str)
            self._validar(clave, df.columns)
            dest = os.path.join(self.rutas.tablas, spec["archivo"])
            df.to_csv(dest, index=False, encoding="utf-8")
            self._registrar(clave, f"{base} [{hoja}]", len(df))
            log(f"{spec['nombre']}: {len(df):,} registros (hoja '{hoja}').")
            hechos += 1
        if not hechos:
            raise ErrorConsolidador("El Excel no contiene ninguna de las hojas esperadas (UPS, Establecimiento, "
                                    "Condicion, Profesion).")

    def vista_previa(self, clave, n=500):
        ruta = self.ruta_maestro(clave)
        if not ruta:
            raise ErrorConsolidador("Este maestro aun no ha sido cargado.")
        return read_csv_safe(ruta, nrows=n).fillna("")

    # --- tramas ---
    def listar_tramas(self):
        archivos = glob.glob(os.path.join(self.rutas.tramas, "**", "*.csv"), recursive=True)
        return sorted(archivos, key=lambda p: os.path.basename(p).lower())

    def agregar_tramas(self, rutas, log=print):
        """Copia CSV o extrae ZIP dentro de TRAMAS. Devuelve rutas agregadas."""
        agregadas = []
        for r in rutas:
            ext = os.path.splitext(r)[1].lower()
            if ext == ".zip":
                log(f"Extrayendo {os.path.basename(r)}...")
                agregadas += extraer_csvs_zip(r, self.rutas.tramas)
            elif ext == ".csv":
                dest = os.path.join(self.rutas.tramas, os.path.basename(r))
                if os.path.abspath(r) != os.path.abspath(dest):
                    shutil.copyfile(r, dest)
                agregadas.append(dest)
            else:
                raise ErrorConsolidador(f"Formato no soportado para tramas: {os.path.basename(r)}")
        return agregadas

    # --- contexto de procesamiento ---
    def cargar_contexto(self, log=print):
        faltan = [MAESTROS[k]["nombre"] for k in ("pacientes", "personal", "registradores", *CLAVES_TABLAS)
                  if not self.ruta_maestro(k)]
        if faltan:
            raise ErrorConsolidador("Faltan datos maestros por cargar: " + ", ".join(faltan))
        t0 = time.time()
        log("Cargando tablas maestras...")
        mapas = {}
        for clave in CLAVES_TABLAS:
            spec = MAESTROS[clave]
            df = read_csv_safe(self.ruta_maestro(clave))
            c_id = obtener_columna(df.columns, spec["id"])
            c_desc = obtener_columna(df.columns, spec["desc"])
            m = {}
            for k, v in zip(df[c_id], df[c_desc]):
                agregar_a_mapa(m, k, v)
            mapas[clave] = m
            log(f"  - {spec['nombre']}: {len(m):,} claves")

        log("Cargando catalogos CIE10 y CPMS...")
        catalogo = {}
        for clave in ("cie10", "cpms"):
            ruta = self.ruta_maestro(clave)
            if not ruta:
                log(f"  ADVERTENCIA: {MAESTROS[clave]['nombre']} no cargado; se omite.")
                continue
            df = read_csv_safe(ruta)
            for cod, desc in zip(df["Codigo_Item"], df["Descripcion_Item"]):
                if pd.notna(cod) and pd.notna(desc):
                    d = str(desc).strip()
                    catalogo[clean_key(cod)] = d
                    catalogo[str(cod).strip()] = d
        log(f"  - Catalogos: {len(catalogo):,} codigos")

        log("Cargando Pacientes, Personal y Registradores...")

        def cargar(clave, indice, renombres):
            cols = MAESTROS[clave]["cols"]
            df = read_csv_safe(self.ruta_maestro(clave), usecols=cols)
            return df.rename(columns=renombres).set_index(indice)

        df_pac = cargar("pacientes", "Id_Paciente", {
            'Id_Tipo_Documento': 'Id_Tipo_Documento_Paciente', 'Numero_Documento': 'Numero_Documento_Paciente',
            'Fecha_Nacimiento': 'Fecha_Nacimiento_Paciente'})
        df_per = cargar("personal", "Id_Personal", {
            'Id_Tipo_Documento': 'Id_Tipo_Documento_Personal', 'Numero_Documento': 'Numero_Documento_Personal',
            'Fecha_Nacimiento': 'Fecha_Nacimiento_Personal'})
        df_reg = cargar("registradores", "Id_Registrador", {
            'Id_Tipo_Documento': 'Id_Tipo_Documento_Registrador',
            'Numero_Documento': 'Numero_Documento_Registrador'})
        log(f"Maestros cargados en {time.time() - t0:.1f}s")
        return dict(ups=mapas["ups"], est=mapas["establecimiento"], cond=mapas["condicion"],
                    prof=mapas["profesion"], catalogo=catalogo, pac=df_pac, per=df_per, reg=df_reg)


# ----------------------------------------------------------------------------
# Procesamiento
# ----------------------------------------------------------------------------

def procesar_trama(ruta_trama, ctx):
    """Devuelve el DataFrame consolidado de una trama (misma logica de procesar_todo.py)."""
    df_raw = read_csv_safe(ruta_trama)
    if df_raw.empty:
        return None
    ups_map, est_map, cond_map, prof_map = ctx["ups"], ctx["est"], ctx["cond"], ctx["prof"]
    catalogo_map = ctx["catalogo"]

    if 'Id_Ups' in df_raw.columns:
        df_raw['Ups'] = df_raw['Id_Ups'].map(lambda x: ups_map.get(clean_key(x), ups_map.get(str(x).strip(), "")))
    else:
        df_raw['Ups'] = ""
    if 'Id_Establecimiento' in df_raw.columns:
        df_raw['Establecimiento'] = df_raw['Id_Establecimiento'].map(
            lambda x: est_map.get(clean_key(x), est_map.get(str(x).strip(), "")))
    else:
        df_raw['Establecimiento'] = ""

    for col in COLUMNAS_FINALES:
        if col not in df_raw.columns:
            df_raw[col] = pd.NA
    df_raw = df_raw[COLUMNAS_FINALES].copy()

    df_raw = df_raw.join(ctx["pac"], on='Id_Paciente', rsuffix='_maestro')
    df_raw = df_raw.join(ctx["per"], on='Id_Personal', rsuffix='_maestro')
    df_raw = df_raw.join(ctx["reg"], on='Id_Registrador', rsuffix='_maestro')

    for col in COLUMNAS_FINALES:
        col_m = f"{col}_maestro"
        if col_m in df_raw.columns:
            df_raw[col] = df_raw[col].fillna(df_raw[col_m])
            df_raw = df_raw.drop(columns=[col_m])

    if 'Id_Condicion' in df_raw.columns:
        df_raw['Condicion'] = df_raw['Id_Condicion'].map(
            lambda x: cond_map.get(clean_key(x), cond_map.get(str(x).strip(), "")))
        df_raw = df_raw.drop(columns=['Id_Condicion'])
    if 'Id_Profesion' in df_raw.columns:
        df_raw['Profesion'] = df_raw['Id_Profesion'].map(
            lambda x: prof_map.get(clean_key(x), prof_map.get(str(x).strip(), "")))
        df_raw = df_raw.drop(columns=['Id_Profesion'])

    def get_descripcion(codigo, actual):
        k = clean_key(codigo)
        desc = catalogo_map.get(k) or catalogo_map.get(str(codigo).strip())
        if desc:
            return desc
        return actual if (pd.notna(actual) and str(actual).strip() != "") else ""

    df_raw['Descripcion_Item'] = [
        get_descripcion(c, d) for c, d in zip(df_raw['Codigo_Item'], df_raw['Descripcion_Item'])
    ]
    return df_raw[COLUMNAS_FINALES]


def nombre_por_defecto(ruta_trama):
    """Nombre del consolidado = mismo nombre de la trama."""
    return os.path.basename(ruta_trama)


def consolidar(gestor, items, modo, nombre_unico="CONSOLIDADO_COMPLETO", excel_bom=False,
               log=print, progreso=lambda f, t: None, cancelar=None):
    """
    items: lista de (ruta_trama, nombre_salida)
    modo: 'separados' | 'unico'
    Devuelve la lista de archivos generados.
    """
    if not items:
        raise ErrorConsolidador("No hay tramas seleccionadas.")
    enc = "utf-8-sig" if excel_bom else "utf-8"
    ctx = gestor.cargar_contexto(log)
    salida = gestor.rutas.salida
    os.makedirs(salida, exist_ok=True)
    t0 = time.time()
    generados = []
    usados = set()

    def unico(nombre):
        base, ext = os.path.splitext(nombre)
        n, cand = 1, nombre
        while cand.lower() in usados:
            n += 1
            cand = f"{base}_{n}{ext}"
        usados.add(cand.lower())
        return cand

    total = len(items)
    ruta_unica = None
    vistos = set()
    escritas = 0
    duplicados = 0
    if modo == "unico":
        ruta_unica = os.path.join(salida, nombre_seguro(nombre_unico))

    try:
        for i, (ruta, nombre) in enumerate(items):
            if cancelar is not None and cancelar.is_set():
                log("Proceso cancelado por el usuario.")
                break
            base = os.path.basename(ruta)
            progreso(i / total, f"Procesando {base} ({i + 1}/{total})")
            log(f"Procesando: {base}")
            t1 = time.time()
            df = procesar_trama(ruta, ctx)
            if df is None:
                log("  -> Archivo vacio, se omite.")
                continue
            if modo == "separados":
                destino = os.path.join(salida, unico(nombre_seguro(nombre)))
                df.to_csv(destino, index=False, encoding=enc)
                generados.append(destino)
                log(f"  -> {os.path.basename(destino)} ({len(df):,} filas) en {time.time() - t1:.1f}s")
            else:
                antes = len(df)
                df = df.drop_duplicates()
                h = pd.util.hash_pandas_object(df, index=False).values
                nuevo = [x not in vistos for x in h]
                vistos.update(h.tolist())
                # un mismo hash repetido dentro del mismo bloque ya fue filtrado por drop_duplicates
                df = df[nuevo]
                duplicados += antes - len(df)
                df.to_csv(ruta_unica, index=False, encoding=enc if escritas == 0 else "utf-8",
                          mode="w" if escritas == 0 else "a", header=(escritas == 0))
                escritas += len(df)
                log(f"  -> agregado al consolidado ({len(df):,} filas nuevas) en {time.time() - t1:.1f}s")
        if modo == "unico" and escritas:
            generados.append(ruta_unica)
            log(f"Consolidado unico: {escritas:,} filas, {duplicados:,} duplicados eliminados.")
    finally:
        progreso(1.0, "Listo")
    log(f"--- PROCESO COMPLETADO EN {time.time() - t0:.1f}s ---")
    return generados
