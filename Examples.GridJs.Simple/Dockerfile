FROM node:24-alpine AS frontend-build
WORKDIR /src

COPY package.json package-lock.json ./
COPY ClientApp/react-gridjs/package.json ClientApp/react-gridjs/package.json
COPY ClientApp/vue-gridjs/package.json ClientApp/vue-gridjs/package.json
COPY ClientApp/angular-gridjs/package.json ClientApp/angular-gridjs/package.json
COPY ClientApp/vanilla-gridjs-npm/package.json ClientApp/vanilla-gridjs-npm/package.json
RUN npm ci

COPY ClientApp ./ClientApp
RUN npm run build

FROM python:3.13-slim AS runtime
WORKDIR /app

RUN apt-get update \
    && apt-get install --yes --no-install-recommends \
        ca-certificates \
        fontconfig \
        fonts-dejavu-core \
        fonts-liberation \
        libgdiplus \
        libicu-dev \
        locales \
    && locale-gen en_US.UTF-8 zh_CN.UTF-8 \
    && rm -rf /var/lib/apt/lists/*

ENV LANG=en_US.UTF-8 \
    LANGUAGE=en_US:en \
    LC_ALL=en_US.UTF-8 \
    DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false \
    GRIDJS_WORKBOOK_DIR=/app/data \
    GRIDJS_CACHE_DIR=/app/cache \
    GRIDJS_WORKBOOK=sample.xlsx \
    PORT=2022

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
COPY --from=frontend-build /src/static/demos ./static/demos

RUN groupadd --gid 1654 gridjs \
    && useradd --uid 1654 --gid gridjs --create-home gridjs \
    && mkdir -p /app/cache \
    && chown -R gridjs:gridjs /app

EXPOSE 2022
VOLUME ["/app/cache"]

USER gridjs

HEALTHCHECK --interval=15s --timeout=5s --start-period=40s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:2022/health', timeout=4).read()" || exit 1

CMD ["gunicorn", "--bind", "0.0.0.0:2022", "--workers", "1", "--threads", "4", "--timeout", "120", "app:app"]
