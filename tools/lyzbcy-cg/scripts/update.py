"""Stage a verified public release after the foreground CG command finishes."""
from datetime import date
from hashlib import sha256
from io import BytesIO
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from urllib.request import Request, urlopen
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
CACHE = Path(os.getenv("LOCALAPPDATA") or Path.home() / ".cache") / "lyzbcy-cg"
REPO = "lyzbcy/JNU-Toolkit"
API = f"https://api.github.com/repos/{REPO}"
REMOTE_VERSION = f"https://raw.githubusercontent.com/{REPO}/main/tools/lyzbcy-cg/VERSION"
USER_AGENT = "lyzbcy-cg-release-updater"
FILES = ("VERSION", "SKILL.md", "scripts/cg.py", "scripts/scan.js", "scripts/update.py")
MAX_SIZE = 1024 * 1024


def version(value):
    if re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+", value):
        return tuple(map(int,value.split(".")))
    raise ValueError("Invalid CG version")


def get(url,limit):
    req = Request(url,headers={"User-Agent":USER_AGENT})
    with urlopen(req,timeout=6) as res:
        body = res.read(limit+1)
    if len(body)>limit:
        raise ValueError("Oversized release response")
    return body


def staged_command():
    """Return the newest completed cache entry, if newer than this copy."""
    try:
        selected = (CACHE / "current").read_text(encoding="ascii").strip()
        if version(selected) <= version((ROOT / "VERSION").read_text(encoding="ascii").strip()):
            return None
        destination = CACHE / selected
        if (destination / "VERSION").read_text(encoding="ascii").strip() == selected:
            return destination / "scripts" / "cg.py"
    except (OSError,ValueError):
        pass
    return None


def schedule():
    args = [sys.executable,str(Path(__file__).resolve())]
    flags = subprocess.CREATE_NO_WINDOW if os.name=="nt" else 0
    try:
        subprocess.Popen(args,stdin=subprocess.DEVNULL,stdout=subprocess.DEVNULL,
                         stderr=subprocess.DEVNULL,creationflags=flags,start_new_session=os.name!="nt")
    except OSError:
        pass


def update():
    try:
        CACHE.mkdir(parents=True,exist_ok=True)
        today=date.today().isoformat()
        checked=CACHE/"checked"
        if checked.exists() and checked.read_text(encoding="ascii").strip()==today:
            return
        checked.write_text(today,encoding="ascii")
        version_local=(ROOT/"VERSION").read_text(encoding="ascii").strip()
        version_remote=get(REMOTE_VERSION,100).decode("ascii").strip()
        if version(version_remote)<=version(version_local):
            return
        tag="cg-v"+version_remote
        release=json.loads(get(f"{API}/releases/tags/{tag}",MAX_SIZE))
        asset=next(a for a in release["assets"] if a["name"]=="lyzbcy-cg.zip")
        url=f"https://github.com/{REPO}/releases/download/{tag}/lyzbcy-cg.zip"
        if release["draft"] or release["prerelease"] or asset["browser_download_url"]!=url:
            return
        blob=get(url,MAX_SIZE)
        if asset.get("digest")!="sha256:"+sha256(blob).hexdigest():
            return
        dest=CACHE/version_remote
        with ZipFile(BytesIO(blob)) as z:
            expected={"lyzbcy-cg/"+f for f in FILES}
            if set(z.namelist())!=expected or sum(x.file_size for x in z.infolist())>MAX_SIZE:
                return
            files={name:z.read("lyzbcy-cg/"+name) for name in FILES}
        if files["VERSION"].decode("ascii").strip()!=version_remote:
            return
        for name,data in files.items():
            path=dest/name
            path.parent.mkdir(parents=True,exist_ok=True)
            pending=path.with_suffix(path.suffix+f".{os.getpid()}.pending")
            pending.write_bytes(data)
            pending.replace(path)
        pending=CACHE/f"current.{os.getpid()}.pending"
        pending.write_text(version_remote,encoding="ascii")
        pending.replace(CACHE/"current")
    except (OSError,ValueError,KeyError,StopIteration):
        return


if __name__=="__main__":
    update()
