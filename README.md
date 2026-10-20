# Sistema Consolidador de Tramas

Aplicación de escritorio (Windows) que homologa tramas nominales (CSV) con los datos maestros
(Pacientes, Registradores, Personal, catálogos CIE10/CPMS y tablas UPS, Establecimiento, Condición y Profesión)
y genera consolidados en CSV.

## Características
- Interfaz tipo formulario con barra de menús superior.
- **Datos maestros**: visualiza cada maestro, su número de registros y la **fecha de última actualización**.
  Se actualizan con **CSV, ZIP o Excel** (ZIP para Pacientes, Registradores y Personal; el CSV se extrae del comprimido).
  Las tablas UPS/Establecimiento/Condición/Profesión pueden cargarse desde el **Excel completo** o con un CSV por tabla.
- **Tramas**: se agregan en **CSV o ZIP**. Puedes procesar varias a la vez.
- **Salida**: consolidados **separados** (nombre editable, por defecto el de la trama) o **un solo consolidado**
  (une todas las tramas y elimina filas idénticas).
- Opción de CSV compatible con Excel (UTF-8 con BOM).

## Carpetas (se crean junto al .exe)
| Carpeta | Contenido |
|---|---|
| `MAESTROS/` | Maestros cargados y `registro_maestros.json` con las fechas |
| `TRAMAS/` | Tramas agregadas |
| `SALIDA/` | Consolidados generados |

## Uso desde código
```bash
pip install -r requirements.txt
python consolidador_app.py
```

## Crear el ejecutable
```bash
pip install pyinstaller
build_exe.bat
```
El resultado queda en `dist/ConsolidadorTramas.exe`.

## Columnas requeridas
- Pacientes: `Id_Paciente, Id_Tipo_Documento, Numero_Documento, Apellido_Paterno_Paciente, Apellido_Materno_Paciente, Nombres_Paciente, Fecha_Nacimiento, Genero, Id_Etnia, Historia_Clinica`
- Personal: `Id_Personal, Id_Tipo_Documento, Numero_Documento, Apellido_Paterno_Personal, Apellido_Materno_Personal, Nombres_Personal, Fecha_Nacimiento, Id_Condicion, Id_Profesion`
- Registradores: `Id_Registrador, Id_Tipo_Documento, Numero_Documento, Apellido_Paterno_Registrador, Apellido_Materno_Registrador, Nombres_Registrador`
- Catálogos: `Codigo_Item, Descripcion_Item`

## ⚠️ Privacidad
Los maestros y tramas contienen **datos personales de salud**. El `.gitignore` excluye `MAESTROS/`, `TRAMAS/`,
`SALIDA/` y todo `*.csv/*.xlsx/*.zip`. **Nunca subas esos archivos a un repositorio público.**

## Licencia
MIT
