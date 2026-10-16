"""Run the VBA tests in headless LibreOffice (VBA compatibility mode).

LibreOffice Basic is not Office VBA, but with `Option VBASupport 1` it is
close enough to catch compile errors and logic bugs without Windows.

Usage:  python3 tools/run_tests.py
Requires LibreOffice and its Python UNO bridge (python3-uno), plus checkouts of
https://github.com/Calmaroon/TextDistanceVBA and
https://github.com/Calmaroon/PhoneticVBA, either next to this repository or
pointed to by TEXTDISTANCEVBA_DIR and PHONETICVBA_DIR.
"""
import glob
import os
import re
import subprocess
import sys
import time

import uno

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PORT = 2084

# (environment variable, repository name, files needed from it)
DEPENDENCIES = [
    ("TEXTDISTANCEVBA_DIR", "TextDistanceVBA",
     ["JaroWinkler.bas", "DamerauLevenshtein.bas", "Phonetic.bas"]),
    ("PHONETICVBA_DIR", "PhoneticVBA",
     ["Phonetic_Encodings/Encode_DoubleMetaphone.bas"]),
]

HARNESS = """Option VBASupport 1
Option Explicit
Public gLog As String
Public Sub LogLine(ByVal Text As String)
    gLog = gLog & Text & Chr(10)
End Sub
Public Function RunTest(ByVal Name As String) As String
    gLog = ""
    Select Case Name
{cases}
    End Select
    RunTest = gLog
End Function
"""

# Worksheet formulas (Calc uses ; as argument separator)
UDF_FORMULAS = [
    ('=ER_MATCHSCORE("Bill Jones";"William Jones")', 0.975),
    ('=ER_MATCHSCORE("John Smith";"Mary Smith")', 0),
    ('=ER_PARSENAME("Smith, John Q.";"LAST")', "SMITH"),
]


def find_dependencies():
    """Paths of the modules this library needs, or exit naming what is missing."""
    parent = os.path.dirname(ROOT)
    paths = []
    for env, repo, files in DEPENDENCIES:
        candidates = [os.environ.get(env, "")]
        candidates += [os.path.join(parent, d) for d in sorted(os.listdir(parent))
                       if d.lower() == repo.lower()]
        found = next((d for d in candidates
                      if d and all(os.path.isfile(os.path.join(d, f)) for f in files)), None)
        if not found:
            sys.exit(f"{repo} not found: clone it next to this repository or set {env}")
        paths += [os.path.join(found, f) for f in files]
    return paths


LIBRARY = find_dependencies() + sorted(glob.glob(os.path.join(ROOT, "*.bas")))
TESTS = sorted(glob.glob(os.path.join(ROOT, "tests", "*.bas")))


def check_udfs(doc):
    sheet = doc.Sheets.getByIndex(0)
    for row, (formula, _) in enumerate(UDF_FORMULAS):
        sheet.getCellByPosition(0, row).Formula = formula
    doc.calculateAll()
    failures = 0
    for row, (formula, expected) in enumerate(UDF_FORMULAS):
        cell = sheet.getCellByPosition(0, row)
        if isinstance(expected, str):
            ok = not cell.getError() and cell.String == expected
        else:
            ok = not cell.getError() and abs(cell.Value - expected) <= 1e-9
        if not ok:
            failures += 1
            print(f"FAIL {formula}: got {cell.String}, expected {expected}")
    print(f"Worksheet formulas: {len(UDF_FORMULAS) - failures} passed, {failures} failed")
    return failures == 0


def enable_macros(ctx):
    """Allow document macros so Calc formulas can call the Basic functions."""
    from com.sun.star.beans import PropertyValue
    provider = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", ctx)
    node = PropertyValue(Name="nodepath",
                         Value="/org.openoffice.Office.Common/Security/Scripting")
    access = provider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
    access.setPropertyValue("MacroSecurityLevel", 0)
    access.commitChanges()
    return (PropertyValue(Name="MacroExecutionMode", Value=4),)  # ALWAYS_EXECUTE_NO_WARN


def check_names():
    """VBA names ignore case: duplicate public names across modules are ambiguous."""
    seen = {}
    ok = True
    for path in LIBRARY:
        with open(path, encoding="cp1252") as f:
            src = f.read()
        found = [re.search(r'VB_Name = "(\w+)"', src).group(1)]
        found += re.findall(r"^(?:Public )?(?:Function|Sub) (\w+)", src, flags=re.M)
        for name in found:
            other = seen.setdefault(name.lower(), os.path.basename(path))
            if other != os.path.basename(path):
                print(f"FAIL duplicate name {name}: {other} and {os.path.basename(path)}")
                ok = False
    return ok


def load_source(path, is_test):
    with open(path, encoding="cp1252") as f:
        src = f.read().replace("\r\n", "\n")
    name = re.search(r'Attribute VB_Name = "(\w+)"', src).group(1)
    src = re.sub(r"^Attribute .*\n", "", src, flags=re.M)
    if is_test:
        src = src.replace("Debug.Print ", "LogLine ")
    return name, "Option VBASupport 1\n" + src


def connect():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    url = f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext"
    for _ in range(60):
        try:
            return resolver.resolve(url)
        except Exception:
            time.sleep(0.5)
    sys.exit("Could not connect to LibreOffice")


def main():
    if not check_names():
        return 1
    office = subprocess.Popen([
        "soffice", "--headless", "--invisible", "--norestore", "--nologo",
        f"--accept=socket,host=localhost,port={PORT};urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        ctx = connect()
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0,
                                           enable_macros(ctx))
        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        lib = libs.getByName("Standard")

        test_subs = []
        for path in LIBRARY + TESTS:
            is_test = path in TESTS
            name, src = load_source(path, is_test)
            lib.insertByName(name, src)
            if is_test:
                test_subs += re.findall(r"^Public Sub (Test_\w+)\(\)", src, flags=re.M)

        cases = "\n".join(f'        Case "{s}": {s}' for s in test_subs)
        lib.insertByName("TestHarness", HARNESS.format(cases=cases))

        provider = doc.getScriptProvider()
        script = provider.getScript(
            "vnd.sun.star.script:Standard.TestHarness.RunTest"
            "?language=Basic&location=document")

        failed = False
        for sub in test_subs:
            output = script.invoke((sub,), (), ())[0]
            print(output if output.strip() else f"{sub}: no output (Basic error?)\n", end="")
            summary = output.strip().splitlines()[-1] if output.strip() else ""
            if not re.search(r" 0 failed$", summary):
                failed = True
        if not check_udfs(doc):
            failed = True
        doc.close(True)
        return 1 if failed or not test_subs else 0
    finally:
        office.terminate()
        office.wait()


if __name__ == "__main__":
    sys.exit(main())
