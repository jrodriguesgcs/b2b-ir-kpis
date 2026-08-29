#!/usr/bin/env python3
"""
Wrap the built dashboard in a client-side password gate for public hosting.

Usage:
    python3 build_locked.py --password "$DASHBOARD_PASSWORD"

The whole dashboard is AES-256-GCM encrypted with a PBKDF2-SHA256 key
(600k iterations). What gets published is a GCS-branded lock screen plus the
ciphertext - without the password the page source carries no readable data.
Rotating the password is: re-run this, redeploy.

Adapted verbatim (mechanism-for-mechanism) from the sibling
`gcs-hubspot-funnel-reporting` dashboard's own `build_locked.py`, since
Vercel's deployment protection does not cover the bare `*.vercel.app`
alias either - the gate here is the real access control, same as there.
"""
import argparse
import base64
import pathlib
import re
import secrets
import sys

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

HERE = pathlib.Path(__file__).resolve().parent
LOGOS = HERE / "assets" / "logos"
ITERS = 600_000

LOCK = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>Institutional Relations KPI Dashboard — Global Citizen Solutions</title>
<link rel="icon" href="data:image/svg+xml;base64,__FAVICON__">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Yrsa:wght@600&family=Heebo:wght@400;500;600&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{min-height:100vh;display:flex;align-items:center;justify-content:center;
  background:#000957;font-family:'Heebo',-apple-system,sans-serif;padding:24px}
.box{width:min(420px,100%);text-align:center}
.sym{width:52px;height:52px;margin:0 auto 22px;display:block}
.sym svg{width:100%;height:100%;display:block}
h1{font-family:'Yrsa',Georgia,serif;font-weight:600;color:#fff;font-size:1.875rem;
  letter-spacing:-.015em;line-height:1.15;margin-bottom:8px}
p.sub{color:rgba(255,255,255,.65);font-size:.875rem;line-height:1.6;margin-bottom:28px}
form{display:flex;gap:8px}
input{flex:1;padding:12px 14px;border-radius:6px;border:1px solid rgba(255,255,255,.25);
  background:rgba(255,255,255,.08);color:#fff;font-size:1rem;font-family:inherit}
input::placeholder{color:rgba(255,255,255,.45)}
input:focus{outline:2px solid #3F8CFF;border-color:transparent}
input:focus:not(:focus-visible){outline:none}
input:focus-visible{outline:2px solid #3F8CFF;border-color:transparent}
button:focus-visible{outline:2px solid #fff;outline-offset:2px}
button{padding:12px 22px;border:none;border-radius:6px;background:#3F8CFF;color:#fff;
  font-weight:600;font-size:.9375rem;font-family:inherit;cursor:pointer}
button:hover{filter:brightness(1.08)}
button:disabled{opacity:.6;cursor:wait}
.err{color:#FF8A8A;font-size:.8125rem;margin-top:14px;min-height:1.3em}
.foot{margin-top:40px;color:rgba(255,255,255,.4);font-size:.75rem;line-height:1.6}
</style>
</head>
<body>
<div class="box">
  <span class="sym">__SYMBOL__</span>
  <h1>Institutional Relations KPI Dashboard</h1>
  <p class="sub">Internal BDM performance dashboard. Commercial data — password required.</p>
  <form id="f">
    <input type="password" id="pw" placeholder="Password" autofocus autocomplete="current-password">
    <button id="go" type="submit">Open</button>
  </form>
  <div class="err" id="err"></div>
  <div class="foot">Global Citizen Solutions &middot; Institutional Relations<br>Data through __THROUGH__</div>
</div>
<script>
const SALT='__SALT__', NONCE='__NONCE__', CT='__CT__', ITERS=__ITERS__;
const un64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
async function open_(pw){
  const mat = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw),
    'PBKDF2', false, ['deriveKey']);
  const key = await crypto.subtle.deriveKey(
    {name:'PBKDF2', salt:un64(SALT), iterations:ITERS, hash:'SHA-256'},
    mat, {name:'AES-GCM', length:256}, false, ['decrypt']);
  const plain = await crypto.subtle.decrypt({name:'AES-GCM', iv:un64(NONCE)}, key, un64(CT));
  const html = new TextDecoder().decode(plain);
  sessionStorage.setItem('b2b_ir_kpi_pw', pw);
  document.open(); document.write(html); document.close();
}
document.getElementById('f').addEventListener('submit', async e => {
  e.preventDefault();
  const btn=document.getElementById('go'), err=document.getElementById('err');
  btn.disabled=true; err.textContent='Decrypting…';
  try { await open_(document.getElementById('pw').value); }
  catch(_){ err.textContent='Wrong password.'; btn.disabled=false;
            document.getElementById('pw').select(); }
});
/* Re-open without re-typing for the rest of the browser tab session. */
const saved = sessionStorage.getItem('b2b_ir_kpi_pw');
if (saved) open_(saved).catch(() => sessionStorage.removeItem('b2b_ir_kpi_pw'));
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--password", required=True)
    ap.add_argument("-i", "--input", default=str(HERE / "../outputs/index.html"))
    ap.add_argument("-o", "--out", default=str(HERE / "../outputs/vercel/index.html"))
    args = ap.parse_args()

    src = pathlib.Path(args.input)
    if not src.exists():
        sys.exit(f"missing {src} - run build_dashboard.py first")
    html = src.read_text()

    salt, nonce = secrets.token_bytes(16), secrets.token_bytes(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32,
                      salt=salt, iterations=ITERS).derive(args.password.encode())
    ct = AESGCM(key).encrypt(nonce, html.encode("utf-8"), None)
    b64 = lambda b: base64.b64encode(b).decode()

    through = (re.search(r'"through":"([\d-]+)"', html) or [None, "—"])[1]
    sym = re.sub(r"<\?xml.*?\?>", "",
                 (LOGOS / "GCS-Symbol-White.svg").read_text(), flags=re.S).strip()

    page = (LOCK
            .replace("__SYMBOL__", sym)
            .replace("__FAVICON__", b64((LOGOS / "GCS-Symbol-Blue.svg").read_bytes()))
            .replace("__THROUGH__", through)
            .replace("__SALT__", b64(salt))
            .replace("__NONCE__", b64(nonce))
            .replace("__CT__", b64(ct))
            .replace("__ITERS__", str(ITERS)))

    out = pathlib.Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(page)
    # Nothing else may sit in the deploy folder - it is served publicly.
    strays = [f.name for f in out.parent.iterdir() if f.name != out.name]
    print(f"wrote {out} ({len(page) // 1024} KB, plaintext {len(html) // 1024} KB)")
    if strays:
        print(f"  WARNING: other files in the deploy folder: {strays}")


if __name__ == "__main__":
    main()
