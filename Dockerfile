# ebimport_splash — Dockerised web UI (Lenex-only)
#
# Build:   docker build -t ebimport-splash:latest .
# Run:     docker run --rm -p 5000:5000 ebimport-splash:latest

FROM python:3.12-slim

RUN apt-get update \
 && apt-get install -y --no-install-recommends tini \
 && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir \
        flask==3.0.3 \
        gunicorn==23.0.0 \
        openpyxl==3.1.5 \
        PyMuPDF==1.25.3

WORKDIR /app
COPY src/ ./src/
COPY scripts/ ./scripts/
COPY webapp ./webapp
COPY template_struct.json ./

ARG BUILD_TIMESTAMP=""
RUN if [ -n "${BUILD_TIMESTAMP}" ]; then echo "${BUILD_TIMESTAMP}" > /app/BUILD_TIMESTAMP; \
    else TZ=America/Toronto date '+%Y-%m-%d %H:%M ET' > /app/BUILD_TIMESTAMP; fi

ENV STAGING_DIR=/tmp/ebimport_staging
ENV PYTHONPATH=/app/src:/app

EXPOSE 5000

ENTRYPOINT ["/usr/bin/tini", "--"]
CMD ["gunicorn", \
     "--bind=0.0.0.0:5000", \
     "--workers=1", \
     "--threads=4", \
     "--timeout=900", \
     "--access-logfile=-", \
     "webapp.app:app"]
