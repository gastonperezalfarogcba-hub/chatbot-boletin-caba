# Chatbot Boletín Oficial CABA — Web V2

Esta versión permite tres formas de uso:

1. **Local / OneDrive sincronizado**: igual que el MVP inicial. Lee una carpeta de tu computadora.
2. **Subida manual de Excel**: sirve para usarlo online sin conectar OneDrive todavía.
3. **OneDrive / SharePoint directo**: pensado para publicarlo como app web y consultar desde cualquier computadora o celular.

La app lee todos los Excel con la estructura del Boletín Oficial CABA, consolida la hoja `Normas`, elimina duplicados y permite consultar por filtros o lenguaje natural.

## Qué cambió respecto del MVP inicial

- Tiene modo web.
- Tiene contraseña simple de acceso (`APP_PASSWORD`).
- Tiene fuente de datos configurable: `local`, `upload` u `onedrive`.
- Tiene conector opcional a Microsoft Graph para OneDrive/SharePoint.
- Permite descargar resultados en CSV o Excel.
- Mejora la interpretación sin IA de frases como `últimos 15 días`, `último mes`, `ayer`, `hoy`, etc.
- Sigue funcionando sin pagar API de OpenAI.

## Estructura esperada del Excel

La app espera una hoja llamada `Normas` con encabezados en la fila 3:

- Fecha
- N° Boletín
- Poder / Organismo
- Tipo de Norma
- Área
- Nombre
- Sumario
- URL Documento
- Tiene Anexos

## Uso local en Windows

1. Descomprimir la carpeta.
2. Hacer doble clic en `run_app.bat`.
3. Abrir `http://localhost:8501` si no se abre automáticamente.

La primera prueba usa el archivo incluido en `ejemplo_boletines`.

## Configuración local

Copiar `.env.example` como `.env` y editar:

```env
DATA_SOURCE=local
BOLETIN_FOLDER=C:\Users\TU_USUARIO\OneDrive - Empresa\Boletines
BOLETIN_DB=./data/boletines.sqlite
APP_PASSWORD=cambiar_esta_clave
```

Luego reiniciar la app.

## Uso diario local

1. Subir el Excel diario a la carpeta configurada.
2. Abrir la app.
3. Tocar **Actualizar base desde la carpeta**.
4. Consultar.

## Uso web sin OneDrive directo

Para publicar rápido y probar desde celular o cualquier computadora:

1. Publicar la app.
2. Configurar `DATA_SOURCE=upload`.
3. Entrar con contraseña.
4. Subir los Excel manualmente desde la pantalla.
5. Tocar **Actualizar base con archivos subidos**.

Es el camino más simple para validar la app online antes de conectar Microsoft Graph.

## Uso web con OneDrive / SharePoint directo

Para que la app lea automáticamente la carpeta compartida de OneDrive/SharePoint, configurar:

```toml
DATA_SOURCE = "onedrive"
APP_PASSWORD = "una_clave_segura"
BOLETIN_DB = "/tmp/boletines.sqlite"

MS_TENANT_ID = "..."
MS_CLIENT_ID = "..."
MS_CLIENT_SECRET = "..."
MS_DRIVE_ID = "..."
MS_FOLDER_PATH = "/Boletines"
MS_RECURSIVE = "false"
```

También se puede configurar por sitio de SharePoint:

```toml
DATA_SOURCE = "onedrive"
APP_PASSWORD = "una_clave_segura"
BOLETIN_DB = "/tmp/boletines.sqlite"

MS_TENANT_ID = "..."
MS_CLIENT_ID = "..."
MS_CLIENT_SECRET = "..."
MS_SITE_HOSTNAME = "empresa.sharepoint.com"
MS_SITE_PATH = "/sites/NombreDelSitio"
MS_FOLDER_PATH = "/Documentos compartidos/Boletines"
```

O por OneDrive de usuario:

```toml
DATA_SOURCE = "onedrive"
APP_PASSWORD = "una_clave_segura"
BOLETIN_DB = "/tmp/boletines.sqlite"

MS_TENANT_ID = "..."
MS_CLIENT_ID = "..."
MS_CLIENT_SECRET = "..."
MS_USER_ID = "usuario@empresa.com"
MS_FOLDER_PATH = "/Boletines"
```

## Activar IA opcional

La app funciona sin IA. Para activar IA:

```toml
OPENAI_API_KEY = "sk-..."
OPENAI_MODEL = "gpt-4o-mini"
```

Si no configurás `OPENAI_API_KEY`, el sistema usa búsqueda por palabras clave y filtros heurísticos.

## Seguridad básica

- No subas `.env` ni `.streamlit/secrets.toml` a GitHub.
- Usá `APP_PASSWORD` si la app queda online.
- Creá una app de Microsoft Graph con permisos de solo lectura.
- Evitá usar cuentas personales para producción.
- Si activás API de OpenAI, cargá la clave como secret, nunca en el código.

## Archivos principales

- `streamlit_app.py`: interfaz web.
- `boletin_core.py`: lectura, limpieza, consolidación y búsquedas.
- `onedrive_graph.py`: conector Microsoft Graph.
- `.env.example`: configuración local.
- `.streamlit/secrets.toml.example`: ejemplo de secrets para nube/local.
- `README_DEPLOY.md`: guía de publicación.


## Cambios V2.1

- Corrige la interpretación de consultas como "Licitaciones de los últimos 15 días".
- Mejora el parser de rangos relativos: días, semanas y meses.
