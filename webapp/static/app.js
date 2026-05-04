// ebimport_splash — tiny front-end, just enough to POST /api/run and
// render the JSON response.

(() => {
  const form       = document.getElementById("run-form");
  const status     = document.getElementById("status");
  const goBtn      = document.getElementById("go");
  const mdbOpts    = document.getElementById("mdb-opts");
  const results    = document.getElementById("results");
  const sumBlock   = document.getElementById("summary-block");
  const sumList    = document.getElementById("summary-list");
  const issBlock   = document.getElementById("issues-block");
  const issList    = document.getElementById("issues-list");
  const dlBlock    = document.getElementById("download-block");
  const dlLink     = document.getElementById("download-link");
  const rawBlock   = document.getElementById("raw-block");
  const rawOutput  = document.getElementById("raw-output");

  // Only show the 'existing mdb' option when mode=mdb
  function syncMdbOpts() {
    const mode = (new FormData(form)).get("mode");
    mdbOpts.style.display = (mode === "mdb" || mode === "dry-run")
      ? "" : "none";
  }
  form.addEventListener("change", (ev) => {
    if (ev.target.name === "mode") syncMdbOpts();
  });
  syncMdbOpts();

  form.addEventListener("submit", async (ev) => {
    ev.preventDefault();
    const mode = (new FormData(form)).get("mode");

    // Reset
    results.hidden = true;
    sumBlock.hidden = issBlock.hidden = dlBlock.hidden = rawBlock.hidden = true;
    sumList.innerHTML = "";
    issList.innerHTML = "";
    rawOutput.textContent = "";
    status.className = "status";
    status.textContent = "Traitement en cours…";
    goBtn.disabled = true;

    try {
      const data = new FormData(form);
      const resp = await fetch("/api/run", { method: "POST", body: data });
      const payload = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        status.className = "status error";
        status.textContent = payload.error || `Erreur HTTP ${resp.status}`;
        goBtn.disabled = false;
        return;
      }

      status.textContent = (payload.returncode === 0)
        ? "Terminé."
        : `Terminé avec code ${payload.returncode}.`;

      results.hidden = false;

      // Summary
      if (payload.summary && payload.summary.length) {
        sumBlock.hidden = false;
        for (const line of payload.summary) {
          const li = document.createElement("li");
          li.textContent = line.replace(/^\+\s*/, "");
          sumList.appendChild(li);
        }
      }

      // Issues
      const issues = payload.issues || {};
      const issueCats = Object.entries(issues);
      if (issueCats.length) {
        issBlock.hidden = false;
        // Sort: WARNINGs first (desc count), then NOTEs (desc count)
        issueCats.sort(([, a], [, b]) => {
          if (a.severity !== b.severity) {
            return a.severity === "WARNING" ? -1 : 1;
          }
          return b.count - a.count;
        });
        for (const [cat, data] of issueCats) {
          const box = document.createElement("div");
          box.className = "cat " + data.severity;
          const hdr = document.createElement("p");
          hdr.className = "cat-header";
          hdr.innerHTML =
            `<span class="sev ${data.severity}">[${data.severity}]</span> ` +
            `<code>${cat}</code> — ${data.count} occurrence${data.count > 1 ? "s" : ""}`;
          box.appendChild(hdr);
          if (data.items && data.items.length) {
            const ul = document.createElement("ul");
            for (const it of data.items) {
              const li = document.createElement("li");
              li.textContent = it.message + (it.row ? ` (ligne ${it.row})` : "");
              ul.appendChild(li);
            }
            box.appendChild(ul);
          }
          issList.appendChild(box);
        }
      } else {
        issBlock.hidden = false;
        issList.innerHTML = "<p><em>Aucun problème détecté.</em></p>";
      }

      // Download
      if (payload.download_id) {
        dlBlock.hidden = false;
        const url = `/api/download/${payload.download_id}` +
                    `?name=${encodeURIComponent(payload.download_name || 'result.zip')}`;
        dlLink.href = url;
        dlLink.textContent = `Télécharger ${payload.download_name || 'result.zip'}`;
      }

      // Raw output (collapsed)
      if (payload.raw_output) {
        rawBlock.hidden = false;
        rawOutput.textContent = payload.raw_output;
      }
    } catch (err) {
      status.className = "status error";
      status.textContent = "Erreur: " + err.message;
    } finally {
      goBtn.disabled = false;
    }
  });
})();
