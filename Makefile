# ---------------------------------------------------------------------------
# ebimport_splash packaging
#
#   make fetch   – download third-party dependencies into build/
#   make dist    – build a self-contained tarball under dist/
#   make clean   – wipe build/ and dist/
#
# The resulting tarball extracts to a directory with run_mdb.sh and
# run_lenex.sh entrypoints that bootstrap a local .venv on first run.
# No internet access is required on the target machine once the tarball
# is built; build-time fetch is the only networked step.
# ---------------------------------------------------------------------------

NAME          := ebimport_splash
VERSION       := 1.0.0
DIST_NAME     := $(NAME)-$(VERSION)

BUILD_DIR     := build
DIST_DIR      := dist
STAGE_DIR     := $(BUILD_DIR)/stage/$(DIST_NAME)

UCANACCESS_VER  := 5.0.1
UCANACCESS_ZIP  := UCanAccess-$(UCANACCESS_VER).bin.zip
# SourceForge serves a meta-refresh interstitial; the Makefile target
# parses it and follows the real download URL.
UCANACCESS_URL  := https://downloads.sourceforge.net/project/ucanaccess/$(UCANACCESS_ZIP)
UCANACCESS_DIR  := $(BUILD_DIR)/ucanaccess

PY_PKGS        := openpyxl jaydebeapi JPype1
PY_WHEELS_DIR  := $(BUILD_DIR)/wheels
# Python versions we support on the target machine.  Wheels for each
# minor version are fetched so the bundled directory satisfies pip
# regardless of which Python the user has installed.
PY_VERSIONS    := 3.10 3.11 3.12 3.13

PYTHON ?= python3

# ---------------------------------------------------------------------------

.PHONY: all fetch dist clean help

all: dist

help:
	@echo "Targets:"
	@echo "  make fetch   – download third-party deps into $(BUILD_DIR)/"
	@echo "  make dist    – build $(DIST_DIR)/$(DIST_NAME).tgz"
	@echo "  make clean   – rm -rf $(BUILD_DIR) $(DIST_DIR)"

# ---------------------------------------------------------------------------
# fetch: populate build/ with UCanAccess jars + Python wheels
# ---------------------------------------------------------------------------

fetch: $(UCANACCESS_DIR)/.done $(PY_WHEELS_DIR)/.done

$(UCANACCESS_DIR)/.done:
	@echo ">> downloading UCanAccess $(UCANACCESS_VER)"
	mkdir -p $(BUILD_DIR)
	# SourceForge serves an HTML interstitial with a meta-refresh to the
	# real file.  Fetch the interstitial, extract the refresh URL, then
	# download the actual zip.
	curl -sSL --retry 3 -A "Mozilla/5.0" \
	    -o $(BUILD_DIR)/_sf_interstitial.html $(UCANACCESS_URL)
	@real_url=$$(grep -oE 'content="[0-9]+; url=[^"]+' \
	    $(BUILD_DIR)/_sf_interstitial.html \
	    | head -1 \
	    | sed -E 's/^.*url=//; s/&amp;/\&/g'); \
	if [ -z "$$real_url" ]; then \
	    echo "ERROR: could not parse SourceForge refresh URL" >&2; \
	    exit 1; \
	fi; \
	echo ">> following to $$real_url"; \
	curl -sSL --retry 3 -A "Mozilla/5.0" \
	    -o $(BUILD_DIR)/$(UCANACCESS_ZIP) "$$real_url"
	@rm -f $(BUILD_DIR)/_sf_interstitial.html
	# Sanity check: must be an actual zip
	@if ! file $(BUILD_DIR)/$(UCANACCESS_ZIP) | grep -q 'Zip archive'; then \
	    echo "ERROR: $(BUILD_DIR)/$(UCANACCESS_ZIP) is not a zip file" >&2; \
	    echo "       got: $$(file $(BUILD_DIR)/$(UCANACCESS_ZIP))" >&2; \
	    echo "       Download manually from" >&2; \
	    echo "       https://sourceforge.net/projects/ucanaccess/files/" >&2; \
	    echo "       and place it at $(BUILD_DIR)/$(UCANACCESS_ZIP)." >&2; \
	    exit 1; \
	fi
	@echo ">> extracting UCanAccess jars"
	rm -rf $(UCANACCESS_DIR)
	mkdir -p $(UCANACCESS_DIR)
	unzip -q -j $(BUILD_DIR)/$(UCANACCESS_ZIP) \
	    'UCanAccess-*/ucanaccess-*.jar' \
	    'UCanAccess-*/lib/jackcess-*.jar' \
	    'UCanAccess-*/lib/hsqldb-*.jar' \
	    'UCanAccess-*/lib/commons-lang3-*.jar' \
	    'UCanAccess-*/lib/commons-logging-*.jar' \
	    -d $(UCANACCESS_DIR)
	@ls -la $(UCANACCESS_DIR)
	touch $@

$(PY_WHEELS_DIR)/.done:
	@echo ">> downloading Python wheels for: $(PY_PKGS)"
	@echo "   (fetching a wheel per Python minor version: $(PY_VERSIONS))"
	mkdir -p $(PY_WHEELS_DIR)
	# Fetch a manylinux x86_64 wheel for each supported Python version
	# so the bundled dir works regardless of which interpreter the user
	# has installed.  Pure-Python packages (openpyxl, jaydebeapi) are
	# version-agnostic, but JPype1 has per-cp wheels and we need all of
	# them.
	@set -e; for v in $(PY_VERSIONS); do \
	    echo "   -- python-version=$$v"; \
	    $(PYTHON) -m pip download \
	        --dest $(PY_WHEELS_DIR) \
	        --python-version $$v \
	        --only-binary=:all: \
	        --platform manylinux2014_x86_64 \
	        --platform manylinux_2_17_x86_64 \
	        $(PY_PKGS) >/dev/null; \
	done
	@ls -la $(PY_WHEELS_DIR) | head -30
	touch $@

# ---------------------------------------------------------------------------
# dist: stage everything into $(STAGE_DIR) and make a tarball
# ---------------------------------------------------------------------------

dist: fetch $(DIST_DIR)/$(DIST_NAME).tgz

$(DIST_DIR)/$(DIST_NAME).tgz:
	@echo ">> staging to $(STAGE_DIR)"
	rm -rf $(STAGE_DIR)
	mkdir -p $(STAGE_DIR)
	# Source files
	cp load_to_mdb.py load_to_lenex.py README.md $(STAGE_DIR)/
	cp -r scripts $(STAGE_DIR)/
	# Vendored deps
	mkdir -p $(STAGE_DIR)/vendor
	cp -r $(UCANACCESS_DIR)  $(STAGE_DIR)/vendor/ucanaccess
	cp -r $(PY_WHEELS_DIR)   $(STAGE_DIR)/vendor/wheels
	# Top-level docs + version stamp
	cp RUNNING.md $(STAGE_DIR)/
	printf '%s\n' '$(DIST_NAME)' > $(STAGE_DIR)/VERSION
	# Make the launchers executable
	chmod +x $(STAGE_DIR)/scripts/*.sh
	@echo ">> creating tarball"
	mkdir -p $(DIST_DIR)
	tar -C $(BUILD_DIR)/stage -czf $@ $(DIST_NAME)
	@ls -la $@

clean:
	rm -rf $(BUILD_DIR) $(DIST_DIR)
