# Aspose.Cells GridJS Simple for Python via .NET — framework examples

[Product Page](https://products.aspose.com/cells/python-net) | [Documentation](https://docs.aspose.com/cells/python-net/aspose-cells-gridjs/) | [API Reference](https://reference.aspose.com/cells/python-net/aspose.cellsgridjs/) | [Support](https://forum.aspose.com/c/cells)

This Flask Simple example matches the .NET Simple client matrix: four npm clients plus one direct browser `<script>` / CDN example.

| Client | Development URL | Production URL | npm import |
|---|---|---|---|
| React | `http://127.0.0.1:5183` | `/demos/react/` | `gridjs-spreadsheet/react` |
| Vue 3 | `http://127.0.0.1:5174` | `/demos/vue/` | `gridjs-spreadsheet/vue` |
| Angular 22 | `http://127.0.0.1:4200` | `/demos/angular/` | `gridjs-spreadsheet/angular` |
| JavaScript | `http://127.0.0.1:5175` | `/demos/npm/` | `gridjs-spreadsheet` |

All four npm clients load `data/sample.xlsx`. Workbook loading and the direct JavaScript example use `/GridJs/*`; the published React, Vue, and Angular adapters use their .NET-compatible `/GridJs2/*` defaults for updates, images, downloads, and OLE objects. Flask exposes both route prefixes. The direct `<script>` example is available at `/legacy`.

## Two JavaScript integration modes

### npm import

The framework-free npm example is in `ClientApp/vanilla-gridjs-npm`:

```javascript
import Spreadsheet from "gridjs-spreadsheet";
import JSZip from "jszip";
import "gridjs-spreadsheet/xspreadsheet.css";

window.JSZip = JSZip;
const spreadsheet = new Spreadsheet(document.getElementById("spreadsheet"), {
  updateMode: "server",
  updateUrl: "/GridJs/UpdateCell",
  mode: "edit"
});
```

Run it together with Flask:

```bash
npm run dev:npm
```

Open <http://127.0.0.1:5175/>.

### Direct `<script>` / CDN import

The original Simple template in `templates/index.html` uses browser scripts and does not require an npm build:

```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.6.0/jszip.min.js"></script>
<link rel="stylesheet" href="https://unpkg.com/gridjs-spreadsheet/xspreadsheet.css">
<script src="https://unpkg.com/gridjs-spreadsheet/xspreadsheet.js"></script>
```

After Flask starts, open <http://127.0.0.1:2022/legacy>.

## Requirements

For local development:

- Python 3.11 or newer
- Node.js 26 (recommended), Node.js 24.15.0+, or Node.js 22.22.3+
- npm

The setup script creates an isolated `.venv` automatically. You do not need to install the Python packages globally.

For Docker, only Docker Desktop or Docker Engine with Compose v2 is required.

## One-command local start

From `Examples.GridJs.Simple`, run:

```bash
npm start
```

The first run:

1. Installs all npm workspace dependencies.
2. Creates `.venv` and installs `requirements.txt`.
3. Builds React, Vue, Angular, and JavaScript npm clients into Flask static resources.
4. Starts Flask on port `2022`.

Open:

<http://127.0.0.1:2022/>

The launcher links to all four built npm examples. Press `Ctrl+C` to stop Flask.

The original Simple page remains available at:

<http://127.0.0.1:2022/legacy>

## Development servers

Start Flask and all four npm development servers:

```bash
npm run dev
```

Open a framework directly using the development URLs in the table above. Vite and Angular proxy both `/GridJs` and `/GridJs2` requests to Flask on port `2022`.

Run only one framework with Flask:

```bash
npm run dev:react
npm run dev:vue
npm run dev:angular
npm run dev:npm
```

If dependencies have not been installed yet, run `npm run setup` first.

## Build and validation

Build all four npm clients:

```bash
npm run build
```

The generated files are written to:

```text
static/demos/
```

Validate the Python source:

```bash
npm run check:python
```

## Configuration

The defaults in `app.py` work both locally and in Docker:

| Environment variable | Default | Purpose |
|---|---|---|
| `GRIDJS_WORKBOOK_DIR` | `./data` | Workbook directory |
| `GRIDJS_WORKBOOK` | `sample.xlsx` | Default workbook |
| `GRIDJS_CACHE_DIR` | `./cache` | GridJS cache |
| `PORT` | `2022` | Flask development-server port |

Example:

```bash
GRIDJS_WORKBOOK_DIR=/data/workbooks npm start
```

## Docker

Build and start the complete Python backend and all four npm frontend demos:

```bash
docker compose up --build
```

The published Aspose.Cells GridJS Python package currently provides a Linux `x86_64` wheel. `compose.yaml` therefore defaults to `linux/amd64`, including on Apple Silicon through Docker Desktop emulation. Override `GRIDJS_PYTHON_PLATFORM` only when a wheel for the target architecture is available.

Open:

<http://localhost:2022/>

Run in the background:

```bash
docker compose up --build --detach
docker compose ps
```

View logs:

```bash
docker compose logs --follow gridjs-python-simple
```

Stop and remove the container:

```bash
docker compose down
```

The GridJS cache is stored in a named volume and survives `docker compose down`. Remove it intentionally with:

```bash
docker compose down --volumes
```

Use another host port:

```bash
GRIDJS_PYTHON_PORT=2023 docker compose up --build --detach
```

Then open <http://localhost:2023/>.

The Compose port binds to `127.0.0.1` by default. Add authentication, request limits, URL validation, and appropriate network controls before exposing GridJS editing endpoints publicly.

## Docker architecture

The Dockerfile uses two stages:

1. Node.js 24 builds React, Vue, Angular, and the framework-free JavaScript npm client.
2. Python 3.13 installs the Python via .NET runtime dependencies and serves Flask through Gunicorn as a non-root user.

The container health endpoint is:

<http://localhost:2022/health>

## Troubleshooting

### Docker Hub token or TLS timeout

Errors such as `failed to fetch anonymous token`, `TLS handshake timeout`, or `EOF` happen before the project build starts and normally indicate Docker Hub connectivity. Retry after the network or proxy is stable:

```bash
docker pull --platform linux/amd64 node:24-alpine
docker pull --platform linux/amd64 python:3.13-slim
docker compose up --build
```

### Framework page returns 404

Run `npm run build` before starting Flask locally, or rebuild the Docker image:

```bash
docker compose up --build
```

### Python package installation fails

Confirm that the active Python version is supported, remove `.venv`, and rerun:

```bash
npm run setup
```

### Workbook cannot be loaded

Confirm that `data/sample.xlsx` exists and that the application can write to `cache`.

### Port 2022 is already in use

Use `GRIDJS_PYTHON_PORT` with Docker, or stop the other local process before starting Flask.

### License

The demo runs in evaluation mode unless you add your Aspose.Cells license-loading code during application startup.
