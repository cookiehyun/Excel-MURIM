"""Run the real VBA game code headless in LibreOffice (VBA compatibility mode).

- src/*.bas are loaded as-is, except Excel/Windows-only modules which are
  replaced by stubs (KeyInput -> scripted input, WinApi -> no-op) or skipped
  (GameLoop, DevTools, Gallery).
- Each script in tools/harness/scripts/ is played tick by tick; output:
  tools/harness/out/<name>.log (per-tick state) and <name>.gif (frames).

usage: python3 tools/harness/run_harness.py [script names...]
Note: this checks game logic, not Excel speed or GetAsyncKeyState.
"""
import os, re, subprocess, sys, time
import uno
from com.sun.star.beans import PropertyValue
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
SRC = os.path.join(ROOT, "src")
OUT = os.path.join(HERE, "out")
SKIP = {"GameLoop", "DevTools", "Gallery", "KeyInput", "WinApi"}
W, H = 192, 100
PIPE = os.environ.get("HARNESS_PIPE", "xlfighter")      # run several harnesses in parallel with different pipes

def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p

def vba_source(path):
    txt = open(path, encoding="ascii").read().replace("\r\n", "\n")
    lines = [l for l in txt.split("\n") if not l.startswith("Attribute VB_")]
    return "Option VBASupport 1\n" + "\n".join(lines)

def keyinput_stub():
    real = open(os.path.join(SRC, "KeyInput.bas"), encoding="ascii").read()
    consts = re.findall(r"^Public Const BTN_\w+ As Long = \d+.*$", real, re.M)
    names = {"L": "BTN_LEFT", "R": "BTN_RIGHT", "U": "BTN_UP", "D": "BTN_DOWN",
             "A": "BTN_LP", "S": "BTN_HP", "Z": "BTN_LK", "X": "BTN_HK", "ESC": "BTN_QUIT",
             "TAB": "BTN_SWAP", "H": "BTN_DBG_HIT", "J": "BTN_DBG_DOWN", "DD": "BTN_DBG_DUMMY",
             "F1": "BTN_DBG_BOX", "C": "BTN_TAG"}
    cases = "\n".join(f'    Case "{k}": mNext({v}) = True' for k, v in names.items())
    return f"""Option VBASupport 1
Option Explicit
' generated stub: scripted input instead of GetAsyncKeyState
{chr(10).join(c.strip() for c in consts)}
Private mCur(0 To 31) As Boolean
Private mPrev(0 To 31) As Boolean
Private mNext(0 To 31) As Boolean
Private mLost As Boolean
Private mBlurNext As Boolean
Public Sub Init()
    Dim i As Long
    For i = 0 To 31
        mCur(i) = False: mPrev(i) = False: mNext(i) = False
    Next i
End Sub
Public Sub ClearNext()
    Dim i As Long
    For i = 0 To 31
        mNext(i) = False
    Next i
End Sub
Public Sub SetNextByName(ByVal nm As String)
    Select Case UCase(nm)
{cases}
    Case "BLUR": mBlurNext = True
    Case Else: Err.Raise 5, "KeyInput", "unknown key " & nm
    End Select
End Sub
Public Sub Poll()
    Dim i As Long
    For i = 0 To 31
        mPrev(i) = mCur(i): mCur(i) = mNext(i)
    Next i
    mLost = mBlurNext: mBlurNext = False
End Sub
Public Function FocusLost() As Boolean
    FocusLost = mLost
End Function
Public Function IsDown(ByVal b As Long) As Boolean
    IsDown = mCur(b)
End Function
Public Function IsPressed(ByVal b As Long) As Boolean
    IsPressed = mCur(b) And Not mPrev(b)
End Function
"""

def modules():
    mods = {}
    for fn in sorted(os.listdir(SRC)):
        if fn.endswith(".bas") and fn[:-4] not in SKIP:
            mods[fn[:-4]] = vba_source(os.path.join(SRC, fn))
    mods["KeyInput"] = keyinput_stub()
    mods["WinApi"] = vba_source(os.path.join(HERE, "WinApiStub.bas"))
    mods["HarnessRunner"] = vba_source(os.path.join(HERE, "HarnessRunner.bas"))
    return mods

class Office:
    def __enter__(self):
        self.proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
                                      "--accept=pipe,name=%s;urp;" % PIPE, "-env:UserInstallation=file:///tmp/lo_profile_%s" % PIPE],
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        ctx = uno.getComponentContext()
        res = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
        for _ in range(80):
            try:
                rctx = res.resolve("uno:pipe,name=%s;urp;StarOffice.ComponentContext" % PIPE); break
            except Exception:
                time.sleep(0.5)
        else:
            raise RuntimeError("soffice did not start")
        self.desk = rctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", rctx)
        return self

    def run(self, script_text, out_path, dump_every):
        doc = self.desk.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
        try:
            libs = doc.BasicLibraries
            if not libs.hasByName("Standard"):
                libs.createLibrary("Standard")
            lib = libs.getByName("Standard")
            for name, src in modules().items():
                lib.insertByName(name, src)
            sp = doc.getScriptProvider()
            scr = sp.getScript("vnd.sun.star.script:Standard.HarnessRunner.RunScript?language=Basic&location=document")
            data_dir = os.path.join(ROOT, "data") + "/"
            return scr.invoke((data_dir, script_text, out_path, dump_every), (), ())[0]
        finally:
            doc.close(True)

    def __exit__(self, *a):
        try:
            self.desk.terminate()
        except Exception:
            pass
        self.proc.wait(timeout=30)

def to_gif(log_path, gif_path, scale=4):
    frames = []
    lines = open(log_path).read().split("\n")
    i = 0
    while i < len(lines):
        if lines[i].startswith("F "):
            img = Image.new("RGB", (W, H)); px = img.load()
            for y in range(H):
                row = lines[i + 1 + y]
                for x in range(W):
                    v = int(row[x * 6:x * 6 + 6], 16)          # BGR long
                    px[x, y] = (v & 255, (v >> 8) & 255, (v >> 16) & 255)
            frames.append(img.resize((W * scale, H * scale), Image.NEAREST))
            i += H + 1
        else:
            i += 1
    if frames:
        frames[0].save(gif_path, save_all=True, append_images=frames[1:], duration=50, loop=0)
    return len(frames)

def main():
    os.makedirs(OUT, exist_ok=True)
    sdir = os.path.join(HERE, "scripts")
    names = sys.argv[1:] or sorted(f[:-4] for f in os.listdir(sdir) if f.endswith(".txt"))
    with Office() as office:
        for name in names:
            script = open(os.path.join(sdir, name + ".txt")).read()
            log = os.path.join(OUT, name + ".log")
            res = office.run(script, log, 1)
            n = to_gif(log, os.path.join(OUT, name + ".gif")) if res.startswith("ok") else 0
            calls = [int(l.split()[2]) for l in open(log) if l.startswith("T ")] if os.path.exists(log) else []
            runs = [int(l.split()[3]) for l in open(log) if l.startswith("T ")] if os.path.exists(log) else []
            stat = ""
            if len(calls) > 1:
                c, r = calls[1:], runs[1:]
                stat = f" | paint calls/frame avg {sum(c)/len(c):.1f} max {max(c)} (runs avg {sum(r)/len(r):.1f} max {max(r)})"
            print(f"{name}: {res} | {n} frames{stat}")

if __name__ == "__main__":
    main()
