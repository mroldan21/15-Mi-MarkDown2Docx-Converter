# MD → DOCX Converter v1

Convierte texto **Markdown** a archivos **Word (.docx)** con estilos nativos, lista para el índice automático de Word.

Stack: **FastAPI + Uvicorn** (backend) · **HTML/CSS/JS Vanilla** (frontend) · **python-docx** (generación del documento) · **Conda** (entorno).

---

## Estructura del proyecto

```
md2docx/
├── app.py             # Backend FastAPI
├── index.html         # Interfaz de usuario (dark mode)
├── environment.yml    # Entorno Conda reproducible
└── README.md
```

---

## Requisitos previos

| Herramienta | Versión mínima | Instalación |
|---|---|---|
| [Conda](https://docs.conda.io/en/latest/miniconda.html) o Miniconda | 23.x | [miniconda.html](https://docs.conda.io/en/latest/miniconda.html) |
| Navegador moderno | — | Chrome, Firefox, Edge, Safari |

No necesitas instalar Python ni pip por separado; Conda los gestiona.

---

## 1 · Crear el entorno Conda

Abre una terminal en la carpeta del proyecto y ejecuta:

```bash
conda env create -f environment.yml
```

Esto instala automáticamente:

- Python 3.11
- FastAPI 0.115
- Uvicorn (con extras `standard` para mejor rendimiento)
- python-docx 1.1
- mistune 3.0 (parser Markdown auxiliar)

### Activar el entorno

```bash
conda activate md2docx
```

> Debes activar el entorno **cada vez** que abras una nueva terminal antes de arrancar el servidor.

### Verificar la instalación (opcional)

```bash
python -c "import fastapi, uvicorn, docx; print('Todo OK')"
```

---

## 2 · Levantar el servidor (backend)

Con el entorno activado, desde la carpeta del proyecto:

```bash
conda activate md2docx
uvicorn app:app --reload --host 0.0.0.0 --port 5000
```

Verás algo como:

```
INFO:     Uvicorn running on http://0.0.0.0:5000 (Press CTRL+C to quit)
INFO:     Started reloader process [xxxxx] using WatchFiles
INFO:     Started server process [xxxxx]
INFO:     Waiting for application startup.
INFO:     Application startup complete.
```

### Flags explicados

| Flag | Qué hace |
|---|---|
| `app:app` | Módulo `app.py` → objeto `app` de FastAPI |
| `--reload` | Reinicia el servidor al guardar cambios (solo desarrollo) |
| `--host 0.0.0.0` | Acepta conexiones desde cualquier interfaz de red |
| `--port 5000` | Puerto de escucha (cámbialo si está ocupado) |

### Alternativa: ejecución directa

También puedes arrancar el servidor así (el bloque `__main__` ya incluye uvicorn):

```bash
python app.py
```

### Verificar que el servidor responde

Abre en el navegador o ejecuta en otra terminal:

```bash
curl http://localhost:5000/
# → {"status":"ok","message":"Markdown → DOCX converter running"}
```

La documentación interactiva de la API está disponible en:

- **Swagger UI** → http://localhost:5000/docs
- **ReDoc** → http://localhost:5000/redoc

---

## 3 · Arrancar la interfaz de usuario

No se necesita ningún servidor adicional. El frontend es un único archivo HTML estático.

### Opción A — Doble clic (más sencillo)

Abre el explorador de archivos, navega hasta la carpeta del proyecto y haz doble clic en `index.html`. El navegador lo abrirá con el protocolo `file://`.

### Opción B — Servidor HTTP local (recomendado en Linux/macOS)

Desde la carpeta del proyecto, con el entorno activado:

```bash
python -m http.server 8080
```

Luego abre en el navegador: http://localhost:8080

> Esta opción evita posibles restricciones de seguridad de algunos navegadores con `file://`.

### Opción C — Live Server (VS Code)

Si usas VS Code, instala la extensión **Live Server** (ritwickdey.LiveServer), haz clic derecho sobre `index.html` y selecciona _Open with Live Server_.

---

## 4 · Uso de la aplicación

1. Escribe o pega tu Markdown en el editor de la izquierda.
2. Pulsa **Convertir a .docx** o usa el atajo `Ctrl+Enter` (`Cmd+Enter` en Mac).
3. El navegador descargará `documento.docx` automáticamente.

### Sintaxis Markdown soportada

| Markdown | Resultado en Word |
|---|---|
| `# Título` | Heading 1 (índice automático) |
| `## Sección` | Heading 2 |
| `### Subsección` | Heading 3 |
| `**negrita**` | Negrita |
| `*cursiva*` | Cursiva |
| `***ambos***` | Negrita + Cursiva |
| `` `código` `` | Fuente monoespaciada (rojo) |
| `- item` o `* item` | Lista con viñetas |
| `1. item` | Lista numerada |
| `> cita` | Cita en bloque (cursiva, sangría) |
| ` ```python ` … ` ``` ` | Bloque de código con resaltado de sintaxis |
| `\| col \| col \|` + fila `\|---\|` | Tabla con cabecera azul y filas alternas |
| `---` | Línea separadora |

### Resaltado de sintaxis en bloques de código

Indica el lenguaje tras las tres comillas invertidas para activar el resaltado:

````
```python
def saludar(nombre: str) -> str:
    return f"Hola, {nombre}!"
```

```javascript
const suma = (a, b) => a + b;
```

```sql
SELECT nombre, edad FROM usuarios WHERE activo = 1;
```
````

Tokens resaltados: **strings** (verde), **palabras clave** (azul), **números** (verde claro), **comentarios** (gris). Funciona con Python, JavaScript, TypeScript, SQL, Bash y otros lenguajes de sintaxis similar.

### Tablas

````
| Nombre   | Rol        | Nivel  |
| :------- | :--------: | -----: |
| Ana      | Backend    | Senior |
| Luis     | Frontend   | Mid    |
````

Los dos puntos en la fila separadora controlan la alineación: `:---` izquierda, `:---:` centrado, `---:` derecha.

> **Índice automático en Word:** porque se usan los estilos nativos `Heading 1/2/3`, puedes insertar un índice en Word con _Referencias → Tabla de contenido_ y se generará automáticamente.

---

## 5 · Cambiar el puerto

Si el puerto 5000 está ocupado, arranca con otro puerto:

```bash
uvicorn app:app --reload --host 0.0.0.0 --port 8000
```

Y actualiza la URL en `index.html` (línea 3 del bloque `<script>`):

```js
const BACKEND_URL = 'http://localhost:8000/convert';
```

---

## 6 · Solución de problemas

| Síntoma | Causa probable | Solución |
|---|---|---|
| `ModuleNotFoundError: No module named 'fastapi'` | Entorno no activado | `conda activate md2docx` |
| `[Errno 48] Address already in use` | Puerto ocupado | Cambiar `--port` o matar el proceso con `kill $(lsof -ti:5000)` |
| `No se puede conectar con el backend` (en el navegador) | Servidor no iniciado | Ejecutar `uvicorn app:app --reload --port 5000` |
| El archivo `.docx` se descarga vacío | Markdown solo tiene espacios | Escribir contenido real en el editor |
| CORS bloqueado en Firefox con `file://` | Restricción del navegador | Usar la **Opción B** (servidor HTTP local) |

---

## 7 · Actualizar dependencias

Para regenerar el entorno desde cero después de modificar `environment.yml`:

```bash
conda env remove -n md2docx
conda env create -f environment.yml
conda activate md2docx
```

---

## Licencia

MIT — libre para uso personal y comercial.
