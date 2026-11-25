"""Best-effort daily documentation update; no private data is sent or read."""
from pathlib import Path
from datetime import date
from urllib.request import Request, urlopen
import base64, hashlib, json, os, re, subprocess, sys

ROOT = Path(__file__).resolve().parents[1]
API = "https://api.github.com/repos/lyzbcy/JNU-Toolkit"
PREFIX = "tools/jnu-graduation-prep/"
MAX = 1024 * 1024

def semver(value):
    if not re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+", value):
        raise ValueError("invalid version")
    return tuple(map(int,value.split(".")))

def read_json(url):
    req=Request(url,headers={"User-Agent":"JNU-graduation-prep-updater","Accept":"application/vnd.github+json"})
    with urlopen(req,timeout=8) as resp:
        raw=resp.read(MAX+1)
    if len(raw)>MAX: raise ValueError("oversized response")
    return json.loads(raw)

def blob(name,commit):
    if name not in ("VERSION","SKILL.md") or not re.fullmatch(r"[a-f0-9]{40}",commit):
        raise ValueError("invalid scope")
    data=read_json(f"{API}/contents/{PREFIX}{name}?ref={commit}")
    if data.get("type")!="file" or data.get("path")!=PREFIX+name:
        raise ValueError("unexpected file")
    raw=base64.b64decode(data["content"],validate=False)
    if len(raw)>MAX or data["size"]!=len(raw): raise ValueError("invalid length")
    actual=hashlib.sha1(b"blob "+str(len(raw)).encode("ascii")+b"\x00"+raw).hexdigest()
    if actual!=data["sha"]: raise ValueError("blob checksum mismatch")
    return raw

def write_atomic(path,data):
    temp=path.with_name(path.name+f".{os.getpid()}.pending")
    temp.write_bytes(data);temp.replace(path)

def validate_skill(raw,version):
    text=raw.decode("utf-8")
    return text.startswith("---\n") and re.search(r"(?m)^name: jnu-graduation-prep$",text) and re.search(r'(?m)^  version: "'+re.escape(version)+r'"$',text)

def update():
    try:
        state=ROOT/".update-check.json"
        today=date.today().isoformat()
        if state.exists() and json.loads(state.read_text(encoding="utf-8")).get("date")==today:
            return
        write_atomic(state,json.dumps({"date":today}).encode())
        local=(ROOT/"VERSION").read_text(encoding="ascii").strip()
        commit=read_json(f"{API}/commits/main")["sha"]
        raw_version=blob("VERSION",commit)
        remote=raw_version.decode("ascii").strip()
        if semver(remote)<=semver(local):return
        doc=blob("SKILL.md",commit)
        if not validate_skill(doc,remote):return
        old=(ROOT/"SKILL.md").read_bytes()
        write_atomic(ROOT/"SKILL.md",doc)
        try:write_atomic(ROOT/"VERSION",raw_version)
        except OSError:
            write_atomic(ROOT/"SKILL.md",old);raise
    except (OSError,ValueError,KeyError,TypeError,UnicodeError):
        return

def background():
    try:
        subprocess.Popen([sys.executable,str(Path(__file__).resolve())],
          stdin=subprocess.DEVNULL,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,
          creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0),start_new_session=os.name!="nt")
    except OSError:pass

if __name__=="__main__":
    if "--background" in sys.argv:background()
    else:update()
