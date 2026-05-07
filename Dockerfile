# ebimport_splash — Dockerised web UI
#
# Build:   docker build -t ebimport-splash:latest .
# Run:     docker run --rm -p 5000:5000 ebimport-splash:latest
#
# The resulting image wraps load_to_mdb.py behind a tiny
# Flask + gunicorn web server.  No persistence: each upload
# lives under /tmp/ebimport_staging/<uuid>/ and is cleaned up when
# the download is streamed or after the 30-min TTL.

FROM python:3.12-slim

# System: Java (for UCanAccess) + unzip (for the build stage)
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
        default-jre-headless \
        unzip \
        curl \
        tini \
 && rm -rf /var/lib/apt/lists/*

# Python runtime dependencies — installed globally in the image since
# the container runs as a dedicated process.
RUN pip install --no-cache-dir \
        flask==3.0.3 \
        gunicorn==23.0.0 \
        openpyxl==3.1.5 \
        jaydebeapi==1.2.3 \
        JPype1==1.5.2 \
        PyMuPDF==1.25.3

# UCanAccess jars (5 of them, flattened into /opt/ucanaccess/)
ARG UCANACCESS_VER=5.0.1
RUN mkdir -p /opt/ucanaccess \
 && curl -sSL -A "Mozilla/5.0" \
        -o /tmp/ucanaccess.html \
        "https://downloads.sourceforge.net/project/ucanaccess/UCanAccess-${UCANACCESS_VER}.bin.zip" \
 && real=$(grep -oE 'content="[0-9]+; url=[^"]+' /tmp/ucanaccess.html \
            | head -1 | sed -E 's/^.*url=//; s/&amp;/\&/g') \
 && curl -sSL -A "Mozilla/5.0" -o /tmp/ucanaccess.zip "$real" \
 && unzip -q -j /tmp/ucanaccess.zip \
        "UCanAccess-*/ucanaccess-*.jar" \
        "UCanAccess-*/lib/jackcess-*.jar" \
        "UCanAccess-*/lib/hsqldb-*.jar" \
        "UCanAccess-*/lib/commons-lang3-*.jar" \
        "UCanAccess-*/lib/commons-logging-*.jar" \
        -d /opt/ucanaccess \
 && rm -f /tmp/ucanaccess.html /tmp/ucanaccess.zip \
 && ls -la /opt/ucanaccess

# App code (loaders + webapp)
WORKDIR /app
COPY load_to_mdb.py load_to_lenex.py copy_prelim_to_masters_final.py audit_pdf.py common.py masters_transfer.vbs masters_transfer.bat template.mdb ./
COPY webapp ./webapp

ARG BUILD_TIMESTAMP=""
RUN if [ -n "${BUILD_TIMESTAMP}" ]; then echo "${BUILD_TIMESTAMP}" > /app/BUILD_TIMESTAMP; \
    else TZ=America/Toronto date '+%Y-%m-%d %H:%M ET' > /app/BUILD_TIMESTAMP; fi

# Point the loader at the bundled UCanAccess jars
ENV UCANACCESS_DIR=/opt/ucanaccess
# Shared staging dir for per-request temp files
ENV STAGING_DIR=/tmp/ebimport_staging
# Flask template resolution
ENV PYTHONPATH=/app

EXPOSE 5000

# tini = tiny init so subprocess cleanup works when gunicorn spawns
# Python loader processes.
ENTRYPOINT ["/usr/bin/tini", "--"]
CMD ["gunicorn", \
     "--bind=0.0.0.0:5000", \
     "--workers=1", \
     "--threads=4", \
     "--timeout=900", \
     "--access-logfile=-", \
     "webapp.app:app"]
