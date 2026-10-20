import pandas as pd
import os
import glob
import time

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
    raise ValueError(f"No se encontro ninguna hoja que coincida con {patrones} entre {sheet_names}")

def obtener_columna(df, candidatos):
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidatos:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    raise ValueError(f"No se encontro ninguna columna coincidente con {candidatos} en {list(df.columns)}")

def cargar_tablas_maestras(excel_path):
    print("Cargando Tablas Maestras desde Excel...")
    t0 = time.time()
    xls = pd.ExcelFile(excel_path)
    sheets = xls.sheet_names
    
    # 1. UPS
    hoja_ups = buscar_hoja(sheets, ["UPS"])
    df_ups = pd.read_excel(xls, hoja_ups, dtype=str)
    c_id_ups = obtener_columna(df_ups, ["id_ups", "cod_ups", "ups"])
    c_desc_ups = obtener_columna(df_ups, ["desc_ups", "descripcion_ups", "ups_desc"])
    ups_map = {}
    for _, r in df_ups.iterrows():
        agregar_a_mapa(ups_map, r[c_id_ups], r[c_desc_ups])

    # 2. ESTABLECIMIENTO
    hoja_est = buscar_hoja(sheets, ["ESTABLECIMIENTO"])
    df_est = pd.read_excel(xls, hoja_est, dtype=str)
    c_id_est = obtener_columna(df_est, ["codRenipress", "id_establecimiento", "cod_establecimiento"])
    c_desc_est = obtener_columna(df_est, ["Establecimiento", "nombre_establecimiento", "desc_establecimiento"])
    est_map = {}
    for _, r in df_est.iterrows():
        agregar_a_mapa(est_map, r[c_id_est], r[c_desc_est])

    # 3. CONDICION
    hoja_cond = buscar_hoja(sheets, ["CONDICION"])
    df_cond = pd.read_excel(xls, hoja_cond, dtype=str)
    c_id_cond = obtener_columna(df_cond, ["id_condicion", "cod_condicion"])
    c_desc_cond = obtener_columna(df_cond, ["desc_condicion", "condicion", "descripcion_condicion"])
    cond_map = {}
    for _, r in df_cond.iterrows():
        agregar_a_mapa(cond_map, r[c_id_cond], r[c_desc_cond])

    # 4. PROFESION
    hoja_prof = buscar_hoja(sheets, ["PROFESION"])
    df_prof = pd.read_excel(xls, hoja_prof, dtype=str)
    c_id_prof = obtener_columna(df_prof, ["id_profesion", "cod_profesion"])
    c_desc_prof = obtener_columna(df_prof, ["Profesion", "desc_profesion", "descripcion_profesion"])
    prof_map = {}
    for _, r in df_prof.iterrows():
        agregar_a_mapa(prof_map, r[c_id_prof], r[c_desc_prof])

    print(f"Tablas maestras cargadas con exito en {time.time() - t0:.2f}s:")
    print(f"  - UPS: {len(ups_map)} registros (Hoja: '{hoja_ups}')")
    print(f"  - Establecimientos: {len(est_map)} registros (Hoja: '{hoja_est}')")
    print(f"  - Condiciones: {len(cond_map)} registros (Hoja: '{hoja_cond}')")
    print(f"  - Profesiones: {len(prof_map)} registros (Hoja: '{hoja_prof}')")

    return ups_map, est_map, cond_map, prof_map

def cargar_catalogos(maestros_dir):
    print("Cargando Catalogos de Descripciones (CIE10 y CPMS)...")
    catalogo_map = {}
    rutas = [os.path.join(maestros_dir, "catalogo_cie10.csv"), os.path.join(maestros_dir, "catalogo_cpms.csv")]
    
    for ruta in rutas:
        if not os.path.exists(ruta):
            continue
        try:
            df = pd.read_csv(ruta, dtype=str, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(ruta, dtype=str, encoding='latin-1')
            
        for _, row in df.iterrows():
            if pd.notna(row.get('Codigo_Item')) and pd.notna(row.get('Descripcion_Item')):
                cod = clean_key(row['Codigo_Item'])
                desc = str(row['Descripcion_Item']).strip()
                catalogo_map[cod] = desc
                catalogo_map[str(row['Codigo_Item']).strip()] = desc
                
    print(f"  - Catalogos CIE10/CPMS cargados: {len(catalogo_map)} codigos mapeados.")
    return catalogo_map

def buscar_csv_en_carpeta(carpeta):
    archivos = glob.glob(os.path.join(carpeta, "**", "*.csv"), recursive=True)
    return archivos[0] if archivos else None

def cargar_maestros_csv(maestros_dir):
    print("Cargando Archivos Maestros (Pacientes, Personal, Registradores)...")
    ruta_pac = buscar_csv_en_carpeta(os.path.join(maestros_dir, "PACIENTES"))
    ruta_per = buscar_csv_en_carpeta(os.path.join(maestros_dir, "PERSONAL"))
    ruta_reg = buscar_csv_en_carpeta(os.path.join(maestros_dir, "REGISTRADORES"))

    def cargar_seguro(ruta, cols):
        if not ruta or not os.path.exists(ruta):
            print(f"ADVERTENCIA: No se encontro maestro en {ruta}")
            return pd.DataFrame(columns=cols)
        try:
            return pd.read_csv(ruta, dtype=str, usecols=cols, encoding='utf-8')
        except UnicodeDecodeError:
            return pd.read_csv(ruta, dtype=str, usecols=cols, encoding='latin-1')

    # Paciente
    cols_paciente = ['Id_Paciente', 'Id_Tipo_Documento', 'Numero_Documento', 'Apellido_Paterno_Paciente',
                     'Apellido_Materno_Paciente', 'Nombres_Paciente', 'Fecha_Nacimiento', 'Genero', 'Id_Etnia', 'Historia_Clinica']
    df_paciente = cargar_seguro(ruta_pac, cols_paciente)
    map_pac = {'Id_Tipo_Documento': 'Id_Tipo_Documento_Paciente', 'Numero_Documento': 'Numero_Documento_Paciente', 'Fecha_Nacimiento': 'Fecha_Nacimiento_Paciente'}
    df_paciente = df_paciente.rename(columns=map_pac).set_index('Id_Paciente')

    # Personal
    cols_personal = ['Id_Personal', 'Id_Tipo_Documento', 'Numero_Documento', 'Apellido_Paterno_Personal',
                     'Apellido_Materno_Personal', 'Nombres_Personal', 'Fecha_Nacimiento', 'Id_Condicion', 'Id_Profesion']
    df_personal = cargar_seguro(ruta_per, cols_personal)
    map_per = {'Id_Tipo_Documento': 'Id_Tipo_Documento_Personal', 'Numero_Documento': 'Numero_Documento_Personal', 'Fecha_Nacimiento': 'Fecha_Nacimiento_Personal'}
    df_personal = df_personal.rename(columns=map_per).set_index('Id_Personal')

    # Registrador
    cols_registrador = ['Id_Registrador', 'Id_Tipo_Documento', 'Numero_Documento', 'Apellido_Paterno_Registrador',
                        'Apellido_Materno_Registrador', 'Nombres_Registrador']
    df_registrador = cargar_seguro(ruta_reg, cols_registrador)
    map_reg = {'Id_Tipo_Documento': 'Id_Tipo_Documento_Registrador', 'Numero_Documento': 'Numero_Documento_Registrador'}
    df_registrador = df_registrador.rename(columns=map_reg).set_index('Id_Registrador')

    return df_paciente, df_personal, df_registrador

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    maestros_dir = os.path.join(base_dir, "MAESTROS")
    tramas_dir = os.path.join(base_dir, "TRAMAS")
    salida_dir = os.path.join(base_dir, "SALIDA")
    os.makedirs(salida_dir, exist_ok=True)
    
    excel_path = os.path.join(maestros_dir, "TablasMaestrasV2.xlsx")

    # 1. Cargar dependencias
    start_time = time.time()
    if not os.path.exists(excel_path):
        print(f"ERROR: No se encontro TablasMaestrasV2.xlsx en {maestros_dir}")
        return
        
    ups_map, est_map, cond_map, prof_map = cargar_tablas_maestras(excel_path)
    catalogo_map = cargar_catalogos(maestros_dir)
    df_paciente, df_personal, df_registrador = cargar_maestros_csv(maestros_dir)

    # Columnas finales estandarizadas
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
        'Apellido_Materno_Paciente', 'Nombres_Paciente', 'Fecha_Nacimiento_Paciente', 'Genero', 'Id_Etnia', 'Historia_Clinica', 'Id_Pais', 
        'Id_Tipo_Documento_Personal', 'Numero_Documento_Personal', 'Apellido_Paterno_Personal', 
        'Apellido_Materno_Personal', 'Nombres_Personal', 'Fecha_Nacimiento_Personal', 'Condicion', 'Profesion', 
        'Id_Tipo_Documento_Registrador', 'Numero_Documento_Registrador', 'Apellido_Paterno_Registrador', 
        'Apellido_Materno_Registrador', 'Nombres_Registrador'
    ]

    # 2. Iterar sobre las tramas y procesar
    archivos_tramas = glob.glob(os.path.join(tramas_dir, "**", "*.csv"), recursive=True)
    if not archivos_tramas:
        print(f"No se encontraron tramas en la carpeta {tramas_dir}")
        return

    print(f"\nSe encontraron {len(archivos_tramas)} tramas para procesar.")
    
    for ruta_trama in archivos_tramas:
        nombre_archivo = os.path.basename(ruta_trama)
        print(f"\nProcesando: {nombre_archivo}")
        t_trama = time.time()
        
        try:
            df_raw = pd.read_csv(ruta_trama, dtype=str, encoding='utf-8')
        except UnicodeDecodeError:
            df_raw = pd.read_csv(ruta_trama, dtype=str, encoding='latin-1')

        if df_raw.empty:
            print("  -> Archivo vacio, omitiendo.")
            continue

        # Homologar UPS y Establecimiento
        if 'Id_Ups' in df_raw.columns:
            df_raw['Ups'] = df_raw['Id_Ups'].map(lambda x: ups_map.get(clean_key(x), ups_map.get(str(x).strip(), "")))
        else:
            df_raw['Ups'] = ""
            
        if 'Id_Establecimiento' in df_raw.columns:
            df_raw['Establecimiento'] = df_raw['Id_Establecimiento'].map(lambda x: est_map.get(clean_key(x), est_map.get(str(x).strip(), "")))
        else:
            df_raw['Establecimiento'] = ""

        # Inicializar columnas que no esten en df_raw
        for col in COLUMNAS_FINALES:
            if col not in df_raw.columns:
                df_raw[col] = pd.NA

        # Preservar solo columnas finales antes del join
        df_raw = df_raw[COLUMNAS_FINALES].copy()

        # Joins con maestros
        df_raw = df_raw.join(df_paciente, on='Id_Paciente', rsuffix='_maestro')
        df_raw = df_raw.join(df_personal, on='Id_Personal', rsuffix='_maestro')
        df_raw = df_raw.join(df_registrador, on='Id_Registrador', rsuffix='_maestro')

        # Aplicar Data de Maestro sobre columnas NA / vacias
        for col in COLUMNAS_FINALES:
            col_m = f"{col}_maestro"
            if col_m in df_raw.columns:
                df_raw[col] = df_raw[col].fillna(df_raw[col_m])
                df_raw = df_raw.drop(columns=[col_m])

        # Mapear Condicion y Profesion desde IDs (traidos por df_personal)
        if 'Id_Condicion' in df_raw.columns:
            df_raw['Condicion'] = df_raw['Id_Condicion'].map(lambda x: cond_map.get(clean_key(x), cond_map.get(str(x).strip(), "")))
            df_raw = df_raw.drop(columns=['Id_Condicion'])
        if 'Id_Profesion' in df_raw.columns:
            df_raw['Profesion'] = df_raw['Id_Profesion'].map(lambda x: prof_map.get(clean_key(x), prof_map.get(str(x).strip(), "")))
            df_raw = df_raw.drop(columns=['Id_Profesion'])

        # Mapear Descripcion Item (CIE10 / CPMS)
        def get_descripcion(codigo, actual):
            k = clean_key(codigo)
            desc = catalogo_map.get(k) or catalogo_map.get(str(codigo).strip())
            if desc:
                return desc
            return actual if (pd.notna(actual) and str(actual).strip() != "") else ""
            
        df_raw['Descripcion_Item'] = [
            get_descripcion(c, d) for c, d in zip(df_raw['Codigo_Item'], df_raw['Descripcion_Item'])
        ]

        # Reordenar al formato final estricto
        df_raw = df_raw[COLUMNAS_FINALES]
        
        ruta_salida = os.path.join(salida_dir, f"CONSOLIDADO_{nombre_archivo}")
        df_raw.to_csv(ruta_salida, index=False, encoding='utf-8')
        print(f"  -> Guardado exitosamente: CONSOLIDADO_{nombre_archivo} ({len(df_raw)} filas) en {time.time() - t_trama:.2f}s")

    print(f"\n========================================================")
    print(f"--- PROCESAMIENTO COMPLETADO EN {time.time() - start_time:.2f}s ---")
    print(f"========================================================")

if __name__ == "__main__":
    main()
