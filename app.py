"""
Cloud Run backend for the Result Portions app.

Design note on multi-user safety: the local desktop/browser version used a
single global "run_state" dict polled by /status — correct for one person on
their own computer, but wrong for a shared website where two courts could
upload at the same moment. This version instead processes each upload
entirely within its own request: a fresh temp directory per request, no
global mutable state, and the finished ZIP returned directly as the
response. Two simultaneous uploads never share any state, regardless of
whether Cloud Run happens to route them to the same or different instances.
"""

import os
import io
import sys
import time
import zipfile
import tempfile
import shutil
from datetime import date

from flask import Flask, request, send_file, jsonify, render_template_string

from extractor_core import process_file, LibreOfficeNotFound
from batch_runner import build_output_docx
from credits import check_credits, deduct_credits, InsufficientCredits, LicenseNotFound, LicenseInactive

app = Flask(__name__)

MAX_FILES = 50  # a practical ceiling per request, comfortably inside Cloud Run's request timeout


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", file=sys.stderr, flush=True)


# --- Firebase Admin init -----------------------------------------------
# Requires a service account key, NOT generated here — get this from
# Firebase Console -> Project Settings -> Service Accounts -> Generate new
# private key, then set it as the GOOGLE_APPLICATION_CREDENTIALS env var
# (a file path) or FIREBASE_SERVICE_ACCOUNT_JSON (the JSON content itself,
# e.g. as a Cloud Run secret) when deploying. Not tested against your real
# project here — no credentials exist in this sandbox to test with; the
# credit logic itself (credits.py) was tested against a mock instead.
_db = None


def get_db():
    global _db
    if _db is not None:
        return _db
    import firebase_admin
    from firebase_admin import credentials, firestore as fb_firestore

    if not firebase_admin._apps:
        cred_json = os.environ.get("FIREBASE_SERVICE_ACCOUNT_JSON")
        if cred_json:
            import json
            cred = credentials.Certificate(json.loads(cred_json))
            firebase_admin.initialize_app(cred)
        else:
            # falls back to GOOGLE_APPLICATION_CREDENTIALS env var (a file
            # path), or Cloud Run's default service account if neither is set
            firebase_admin.initialize_app()
    _db = fb_firestore.client()
    return _db


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/api/process", methods=["POST"])
def api_process():
    log("=== new request received ===")
    mode = int(request.form.get("mode", 1))
    uploaded = request.files.getlist("files")
    log(f"mode={mode}, files received={len(uploaded)}")

    if not uploaded:
        return jsonify({"error": "No files were uploaded."}), 400
    if len(uploaded) > MAX_FILES:
        return jsonify({"error": f"Too many files in one batch (max {MAX_FILES}). Please split into smaller batches."}), 400

    workdir = tempfile.mkdtemp(prefix="job_")
    log(f"workdir={workdir}")
    try:
        saved_paths = []
        for f in uploaded:
            ext = os.path.splitext(f.filename)[1].lower()
            if ext not in (".doc", ".docx", ".odt", ".pdf"):
                log(f"skipping unsupported file: {f.filename}")
                continue
            safe_path = os.path.join(workdir, os.path.basename(f.filename))
            log(f"saving {f.filename} -> {safe_path} ...")
            f.save(safe_path)
            log(f"saved OK: {f.filename}")
            saved_paths.append(safe_path)

        log(f"all files saved, count={len(saved_paths)}")

        if not saved_paths:
            return jsonify({"error": "No supported files (.doc, .docx, .odt, .pdf) were found in the upload."}), 400

        license_key = request.form.get("license_key", "").strip()
        if not license_key:
            return jsonify({"error": "Missing license key."}), 400

        log(f"checking credits for {license_key}, {len(saved_paths)} file(s)...")
        try:
            check_credits(get_db(), license_key, len(saved_paths))
        except LicenseNotFound as e:
            return jsonify({"error": str(e)}), 403
        except LicenseInactive as e:
            return jsonify({"error": str(e)}), 403
        except InsufficientCredits as e:
            return jsonify({"error": str(e)}), 402
        log("credits check passed (not yet deducted — happens after the batch succeeds)")

        # upfront LibreOffice check, same as the desktop version — one clear
        # error instead of the same failure repeated per file
        needs_libreoffice = any(p.lower().endswith((".doc", ".odt")) for p in saved_paths)
        if needs_libreoffice:
            log("checking for LibreOffice...")
            from extractor_core import find_soffice
            find_soffice()
            log("LibreOffice found OK")

        groups = {}
        seen_case_numbers = {}

        def get_group(case_type):
            return groups.setdefault(case_type, {"entries": [], "flags": []})

        for path in saved_paths:
            fname = os.path.basename(path)
            log(f"processing {fname} ...")
            t0 = time.time()
            result = process_file(path, mode=mode, workdir=workdir)
            log(f"done {fname} in {round(time.time()-t0, 2)}s")

            if "error" in result:
                g = get_group(result.get("case_type", "UNCLASSIFIED"))
                g["flags"].append(f'Could not process: {fname} — {result["error"]}. Skipped — please check this file manually.')
                continue

            norm_case_no = result["case_no"].replace(" ", "").upper()
            if norm_case_no in seen_case_numbers:
                g = get_group(result["case_type"])
                g["flags"].append(
                    f'Duplicate: {result["case_no"]} already processed from '
                    f'{seen_case_numbers[norm_case_no]} — this entry from {fname} skipped.'
                )
                continue
            seen_case_numbers[norm_case_no] = fname
            get_group(result["case_type"])["entries"].append(result)

        log("all files processed, building output...")
        if not groups:
            return jsonify({"error": "No files could be processed."}), 400

        today = date.today()
        zip_buffer = io.BytesIO()
        written_names = []
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for case_type, group in groups.items():
                if not group["entries"] and not group["flags"]:
                    continue
                out_path = build_output_docx(case_type, group, workdir, today)
                zf.write(out_path, os.path.basename(out_path))
                written_names.append(os.path.basename(out_path))

        # Only now, with the batch genuinely finished and the ZIP ready, is
        # the credit actually deducted — this is the whole reason
        # check_credits and deduct_credits are separate calls (see credits.py).
        # A deduction failure here doesn't block the download (a rare billing
        # hiccup shouldn't cost someone their finished result) but IS logged
        # clearly so it can be reconciled manually.
        try:
            new_remaining = deduct_credits(get_db(), license_key, len(saved_paths))
            log(f"credits deducted, {new_remaining} remaining")
        except Exception as e:
            log(f"WARNING: credit deduction failed after successful processing: {e}")

        zip_buffer.seek(0)
        response = send_file(
            zip_buffer, mimetype="application/zip", as_attachment=True,
            download_name=f"Result files {today.day}-{today.month}-{today.year}.zip",
        )
        response.headers["X-Files-Written"] = ", ".join(written_names)
        return response

    except LibreOfficeNotFound as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Unexpected error: {e}"}), 500
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


@app.route("/")
def index():
    return render_template_string(PAGE)


PAGE = """<!DOCTYPE html>
<html>
<head>
<script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-9117928033972031"
     crossorigin="anonymous"></script>
<meta charset="utf-8">
<title>Result Portions</title>
<style>
  :root { --green:#2e7d32; --green-bg:#e8f5e9; --blue:#1565c0; --blue-bg:#e3f2fd;
          --red:#c62828; --red-bg:#fdecea; --grey:#616161; --grey-bg:#f0f0f0; --border:#dcdcdc; }
  body { font-family: Arial, sans-serif; max-width: 780px; margin: 0 auto; padding: 30px 24px 60px;
         color:#222; background:#fafafa; }
  .card { background:white; border:1px solid var(--border); border-radius:10px; padding:24px 28px;
          margin-bottom:20px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }
  h1 { font-size:21px; margin:0 0 4px; }
  .subtitle { color:#666; font-size:13px; margin:0 0 20px; }
  label { display:block; margin-top:16px; margin-bottom:6px; font-weight:bold; font-size:14px; }
  input[type=file] { width:100%; padding:9px 10px; border:1px solid var(--border); border-radius:6px;
                      font-size:14px; background:#fcfcfc; }
  .mode-card { border:1px solid var(--border); border-radius:8px; padding:12px 14px; margin-top:8px;
               cursor:pointer; display:flex; gap:10px; align-items:flex-start; }
  .mode-card:hover { background:#f7f9fc; }
  .mode-title { font-weight:bold; font-size:14px; }
  .mode-desc { font-size:12.5px; color:#666; margin-top:2px; }
  #runBtn { margin-top:22px; padding:12px 28px; background:var(--green); color:white; border:none;
            border-radius:6px; font-size:15px; font-weight:bold; cursor:pointer; }
  #runBtn:disabled { background:#a5a5a5; cursor:default; }
  .status-badge { display:inline-flex; align-items:center; gap:8px; padding:8px 16px; border-radius:20px;
                  font-weight:bold; font-size:14px; margin-top:18px; }
  .status-idle { background:var(--grey-bg); color:var(--grey); }
  .status-started { background:var(--blue-bg); color:var(--blue); }
  .status-concluded { background:var(--green-bg); color:var(--green); }
  .status-obstructed { background:var(--red-bg); color:var(--red); }
  .dot { width:9px; height:9px; border-radius:50%; background:currentColor; }
  .spin { animation: spin 1s linear infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }
  #messageBox { font-size:14px; margin-top:10px; padding:12px 14px; border-radius:6px; display:none; }
  #messageBox.show-error { display:block; background:var(--red-bg); color:var(--red); border:1px solid #f5c6c2; }
  #messageBox.show-ok { display:block; background:var(--green-bg); color:var(--green); border:1px solid #c3e6c5; }
  #licenseGate input[type=text] { width:100%; padding:9px 10px; border:1px solid var(--border); border-radius:6px; font-size:14px; }
  #licenseGate button { margin-top:14px; padding:10px 22px; background:var(--green); color:white; border:none;
                         border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer; }
  #licenseError { color:var(--red); font-size:13px; margin-top:8px; display:none; }
  #mainApp { display:none; }
</style>
</head>
<body>

  <div class="card" id="licenseGate">
    <h1>Result Portions</h1>
    <p class="subtitle">Enter your license key to continue.</p>
    <label>License Key</label>
    <input type="text" id="licenseKeyInput" placeholder="e.g. COURT001A">
    <button onclick="submitLicenseKey()">Unlock</button>
    <div id="licenseError"></div>
  </div>

  <div class="card" id="mainApp">
    <h1>Result Portions</h1>
    <p class="subtitle">Upload case files — .doc, .docx, .odt, or .pdf (not scanned).</p>

    <label>Case files</label>
    <input type="file" id="fileInput" multiple accept=".doc,.docx,.odt,.pdf">

    <label>Extraction mode</label>
    <div class="mode-card" onclick="document.getElementById('mode1').checked=true">
      <input type="radio" name="mode" value="1" id="mode1" checked>
      <div><div class="mode-title">Decreetal Order Result</div>
      <div class="mode-desc">Decree / Decretal Order, falling back to Order. Judgment is never used.</div></div>
    </div>
    <div class="mode-card" onclick="document.getElementById('mode2').checked=true">
      <input type="radio" name="mode" value="2" id="mode2">
      <div><div class="mode-title">Final Result Portion in Judgment</div>
      <div class="mode-desc">The paragraph beginning "In the result..."</div></div>
    </div>

    <button id="runBtn" onclick="runBatch()">Run</button>
    <div><span id="statusBadge" class="status-badge status-idle"><span class="dot"></span>Ready</span></div>
    <div id="messageBox"></div>
  </div>

<script type="module">
  // License check — same pattern as Pendency Report / 164 Statement:
  // key checked against Firestore, session cached locally with expiry.
  // TODO: replace with the actual config from your other two apps
  // (Firebase Console -> Project Settings -> your existing web app).
  import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
  import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

  const FIREBASE_CONFIG = {
    apiKey: "AIzaSyAVCpa5Pji_G9XMR0CnbcV82Mpfbp6Ou9Y",
    authDomain: "pendency-report-7669a.firebaseapp.com",
    projectId: "pendency-report-7669a",
    storageBucket: "pendency-report-7669a.appspot.com",
    messagingSenderId: "402619057383",
    appId: "REPLACE_ME",
  };
  const db = getFirestore(initializeApp(FIREBASE_CONFIG));

  const SESSION_KEY = "resultportions_session";
  const SESSION_DAYS = 3;

  function parseExpiry(expiryStr) {
    const [d, m, y] = expiryStr.split("-").map(Number);
    return new Date(y, m - 1, d);
  }

  function getLocalSession() {
    try {
      const session = JSON.parse(localStorage.getItem(SESSION_KEY));
      if (!session || new Date(session.validUntil) < new Date()) return null;
      return session;
    } catch { return null; }
  }

  function saveLocalSession(licenseKey) {
    const validUntil = new Date();
    validUntil.setDate(validUntil.getDate() + SESSION_DAYS);
    localStorage.setItem(SESSION_KEY, JSON.stringify({ licenseKey, validUntil: validUntil.toISOString() }));
  }

  function unlockApp() {
    document.getElementById('licenseGate').style.display = 'none';
    document.getElementById('mainApp').style.display = 'block';
  }

  window.getActiveLicenseKey = function () {
    const session = getLocalSession();
    return session ? session.licenseKey : null;
  };

  window.submitLicenseKey = async function () {
    const key = document.getElementById('licenseKeyInput').value.trim();
    const errEl = document.getElementById('licenseError');
    errEl.style.display = 'none';
    if (!key) { errEl.textContent = 'Please enter a license key.'; errEl.style.display = 'block'; return; }
    try {
      const snap = await getDoc(doc(db, 'licenses_resultportions', key));
      if (!snap.exists()) throw new Error('Key not found.');
      const data = snap.data();
      if (data.active === false) throw new Error('This key has been deactivated.');
      if (data.expiryDate && parseExpiry(data.expiryDate) < new Date()) throw new Error('This key has expired.');
      saveLocalSession(key);
      unlockApp();
    } catch (e) {
      errEl.textContent = e.message || 'Could not verify key.';
      errEl.style.display = 'block';
    }
  };

  if (getLocalSession()) unlockApp();
</script>

<script>
function setBadge(status, text, spinning) {
  const el = document.getElementById('statusBadge');
  el.className = 'status-badge status-' + status;
  el.innerHTML = '<span class="dot' + (spinning ? ' spin' : '') + '"></span>' + text;
}
function showMessage(kind, text) {
  const el = document.getElementById('messageBox');
  el.className = kind === 'error' ? 'show-error' : (kind === 'ok' ? 'show-ok' : '');
  el.textContent = text;
}

async function runBatch() {
  const files = document.getElementById('fileInput').files;
  const mode = document.querySelector('input[name=mode]:checked').value;
  const btn = document.getElementById('runBtn');
  showMessage('', '');

  if (!files.length) { showMessage('error', 'Please choose one or more files first.'); return; }

  const licenseKey = window.getActiveLicenseKey ? window.getActiveLicenseKey() : null;
  if (!licenseKey) { showMessage('error', 'Your session expired — please enter your license key again.'); return; }

  const formData = new FormData();
  for (const f of files) formData.append('files', f);
  formData.append('mode', mode);
  formData.append('license_key', licenseKey);

  btn.disabled = true;
  setBadge('started', 'Started — processing…', true);

  try {
    const resp = await fetch('/api/process', { method: 'POST', body: formData });
    if (!resp.ok) {
      const data = await resp.json();
      setBadge('obstructed', 'Obstructed', false);
      showMessage('error', data.error || 'Something went wrong.');
      return;
    }
    const blob = await resp.blob();
    const writtenHeader = resp.headers.get('X-Files-Written') || '';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Result files.zip';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);

    setBadge('concluded', 'Concluded', false);
    showMessage('ok', 'Done — downloaded: ' + (writtenHeader || 'Result files.zip'));
  } catch (e) {
    setBadge('obstructed', 'Obstructed', false);
    showMessage('error', 'Unexpected error: ' + e.message);
  } finally {
    btn.disabled = false;
  }
}
</script>
</body>
</html>"""


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
