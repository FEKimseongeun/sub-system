import re

RE_LINENO = re.compile(r"^\d{1,4}[A-Z0-9\-]+-[A-Z0-9\-]+")   # 하이픈 2개+ 권장
RE_VALVE  = re.compile(r"^(GV|CV|XV|HV)-?\d+[A-Z0-9\-]*$")
RE_SPEC   = re.compile(r"^(SPV|EXJ)[\s\-_]*\d+[A-Z0-9\-]*$")
RE_INSTR  = re.compile(r"^(PT|TT|FT|LT|PI|TI|FI|LG|LIC|TIC|FIC)[A-Z\-]*\d+[A-Z0-9\-]*$")

def classify(text: str) -> str:
    t = text.strip().replace("—","-").replace("–","-")
    if RE_VALVE.match(t):  return "Valve"
    if RE_SPEC.match(t):   return "Special"
    if RE_INSTR.match(t):  return "Instrument"
    if t.count("-") >= 2 and RE_LINENO.match(t): return "LineNo"
    return "Unknown"
