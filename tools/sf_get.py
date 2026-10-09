import glob, os, re, subprocess, sys, urllib.parse
probe, dest = sys.argv[1], sys.argv[2]
for f in glob.glob(os.path.join(probe, "list_*.html")):
    enc_dir = os.path.basename(f)[5:-5]
    name = urllib.parse.unquote(enc_dir)
    os.makedirs(os.path.join(dest, name), exist_ok=True)
    for fn in re.findall(r'<tr title="([^"]+)" class="file', open(f, encoding="utf-8").read()):
        url = "https://downloads.sourceforge.net/project/bidbase/" + enc_dir + "/" + urllib.parse.quote(fn)
        out = os.path.join(dest, name, fn)
        subprocess.run(["curl", "-sSL", "-m", "300", "-o", out, url], check=True)
        print(os.path.getsize(out), name, "/", fn)
