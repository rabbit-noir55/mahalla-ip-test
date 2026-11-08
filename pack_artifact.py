"""Artifactni parol bilan shifrlab paketlaydi (public repo uchun)."""
import glob
import os

import pyzipper

pwd = os.environ["ARTIFACT_PASS"].encode()
files = []
for p in ["/tmp/report.jsonl", "/tmp/botout.log"] + glob.glob("state_v2_*.json") + glob.glob("run_new_*.log"):
    if os.path.exists(p):
        files.append(p)

with pyzipper.AESZipFile("/tmp/report_backup.zip", "w",
                         compression=pyzipper.ZIP_DEFLATED,
                         encryption=pyzipper.WZ_AES) as z:
    z.setpassword(pwd)
    for f in files:
        z.write(f, os.path.basename(f))
print("packed", len(files), "files")
