#!/usr/bin/env python3
"""
Probe VBA class-module support in LibreOffice, over UNO.

Runs *inside* a Python that can import `uno` (LibreOffice's own interpreter,
or system python3 with python3-uno installed). Connects to an soffice
instance already listening on a socket, opens the workbook, and invokes each
VBA test function individually so one failure does not hide the rest.

Two things about LibreOffice that this script exists to work around:

1. MacroExecMode.ALWAYS_EXECUTE_NO_WARN is 9, not 4. 4 is
   USE_CONFIG_REJECT_CONFIRMATION, which silently refuses to run anything.

2. Functions in the document's VBAProject library cannot be invoked through
   the Basic script provider at all -- getScript() resolves them, invoke()
   returns the type's default value, and no code runs. Verified with a
   trivial `Function VbEcho() As String : VbEcho = "x" : End Function`
   inserted straight into VBAProject: still empty. The same function in the
   document's Standard library runs fine. So we inject a Standard-library
   bridge module whose wrappers call VBAProject.<Module>.<Function>(), and
   invoke the wrappers instead.

Usage:
    python3 uno_class_probe.py <workbook.xlsm> <result.json> [port]
"""

import json
import sys
import traceback

# (group, module, function). Each VBA function returns "PASS: ..." / "FAIL: ...".
TESTS = [
    ("basic", "ClassBasicTests", "TestInstantiate"),
    ("basic", "ClassBasicTests", "TestDimAsNew"),
    ("basic", "ClassBasicTests", "TestPropertyGetLet"),
    ("basic", "ClassBasicTests", "TestPropertySet"),
    ("basic", "ClassBasicTests", "TestPrivateState"),
    ("basic", "ClassBasicTests", "TestClassInitialize"),
    ("basic", "ClassBasicTests", "TestTypeName"),
    ("basic", "ClassBasicTests", "TestCollectionOfObjects"),
    ("lifecycle", "ClassLifecycleTests", "TestInitializeCounter"),
    ("lifecycle", "ClassLifecycleTests", "TestTerminateCounter"),
    ("predeclared", "ClassPredeclaredTests", "TestPredeclaredInstance"),
    ("predeclared", "ClassPredeclaredTests", "TestPredeclaredIsShared"),
    ("implements", "ClassImplementsTests", "TestImplementsCompiles"),
    ("implements", "ClassImplementsTests", "TestInterfacePolymorphism"),
    ("events", "ClassEventsTests", "TestEventClassesCompile"),
    ("events", "ClassEventsTests", "TestRaiseEventDelivered"),
]

BRIDGE_MODULE = "ZProbeBridge"
VBA_LIBRARY = "VBAProject"

# MacroExecMode.ALWAYS_EXECUTE_NO_WARN
MACRO_EXEC_ALWAYS_NO_WARN = 9


def bridge_source(tests):
    """
    Build a plain-StarBasic module of wrappers, one per VBA test function.

    Plain Basic (no `Option VBASupport 1`) is deliberate: this module only has
    to reach across into the VBA library and trap errors. `Err` / `Error$` here
    are StarBasic's numeric error code and message, not VBA's Err object.
    """
    parts = []
    for _group, module, func in tests:
        parts.append(
            f"Function W_{func}() As String\n"
            f"    Dim r As String\n"
            f"    On Error Resume Next\n"
            f"    r = {VBA_LIBRARY}.{module}.{func}()\n"
            f"    If Err <> 0 Then\n"
            f'        W_{func} = "BASIC_ERROR: " & Err & " - " & Error$\n'
            f"    Else\n"
            f"        W_{func} = r\n"
            f"    End If\n"
            f"End Function\n"
        )
    return "\n".join(parts)


def read_vba_config(ctx, report):
    """Report the Load/Save > VBA Properties settings that gate VBA import."""
    try:
        from com.sun.star.beans import PropertyValue

        provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx
        )
        arg = PropertyValue()
        arg.Name = "nodepath"
        arg.Value = "/org.openoffice.Office.Calc/Filter/Import/VBA"
        access = provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess", (arg,)
        )
        settings = {
            name: access.getPropertyValue(name) for name in access.getElementNames()
        }
        report.append(f"VBA import config: {settings}")
        return settings
    except Exception as exc:
        report.append(f"could not read VBA config: {type(exc).__name__} {exc}")
        return {}


def describe_libraries(doc, report):
    """List the Basic libraries and modules LibreOffice built from the VBA project."""
    found = {}
    try:
        libs = doc.BasicLibraries
    except Exception as exc:
        report.append(f"BasicLibraries unavailable: {type(exc).__name__} {exc}")
        return found

    for lib_name in libs.getElementNames():
        try:
            if not libs.isLibraryLoaded(lib_name):
                libs.loadLibrary(lib_name)
            lib = libs.getByName(lib_name)
            modules = list(lib.getElementNames())
            found[lib_name] = modules
            report.append(f"library {lib_name}: {len(modules)} modules")

            # LibreOffice stamps imported VBA with these markers. They are the
            # evidence that it recognised our pyOpenVBA-written .cls modules as
            # class modules rather than plain standard modules.
            for mod in sorted(modules):
                try:
                    code = lib.getByName(mod)
                except Exception:
                    continue
                marks = []
                if "Option VBASupport 1" in code:
                    marks.append("VBASupport")
                if "Option ClassModule" in code:
                    marks.append("ClassModule")
                if "VBA_ModuleType=VBAClassModule" in code:
                    marks.append("type:VBAClassModule")
                elif "VBA_ModuleType=VBAModule" in code:
                    marks.append("type:VBAModule")
                report.append(f"  {mod}: {', '.join(marks) or 'no VBA markers'}")
        except Exception as exc:
            report.append(f"library {lib_name}: {type(exc).__name__} {exc}")

    return found


def install_bridge(doc, report):
    """Insert the Standard-library bridge module into the open document."""
    libs = doc.BasicLibraries
    if not libs.hasByName("Standard"):
        libs.createLibrary("Standard")
    if not libs.isLibraryLoaded("Standard"):
        libs.loadLibrary("Standard")
    std = libs.getByName("Standard")

    source = bridge_source(TESTS)
    if std.hasByName(BRIDGE_MODULE):
        std.replaceByName(BRIDGE_MODULE, source)
    else:
        std.insertByName(BRIDGE_MODULE, source)
    report.append(f"bridge module Standard.{BRIDGE_MODULE} installed ({len(TESTS)} wrappers)")


def invoke_wrapper(msp, func):
    """Invoke one bridge wrapper; return (status, detail)."""
    url = (
        f"vnd.sun.star.script:Standard.{BRIDGE_MODULE}.W_{func}"
        "?language=Basic&location=document"
    )
    try:
        script = msp.getScript(url)
    except Exception as exc:
        return "NOT_FOUND", f"{type(exc).__name__}: {str(exc)[:200]}"

    try:
        ret = script.invoke((), (), ())
    except Exception as exc:
        return "ERROR", f"{type(exc).__name__}: {str(exc)[:200]}"

    value = ret[0] if ret else None
    text = "" if value is None else str(value)

    if text.startswith("PASS"):
        return "PASS", text
    if text.startswith("FAIL"):
        return "FAIL", text
    if text.startswith("BASIC_ERROR"):
        return "ERROR", text
    if not text:
        return "NO_RESULT", "empty return -- the VBA function did not run to completion"
    return "UNKNOWN", text


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 2

    workbook = sys.argv[1]
    result_path = sys.argv[2]
    port = sys.argv[3] if len(sys.argv) > 3 else "2002"

    out = {"connected": False, "report": [], "results": [], "libraries": {}}
    report = out["report"]

    try:
        import uno
        from com.sun.star.beans import PropertyValue
    except ImportError as exc:
        report.append(f"cannot import uno: {exc}")
        report.append("Run this under LibreOffice's python or install python3-uno.")
        _write(result_path, out)
        return 1

    try:
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )
        ctx = resolver.resolve(
            f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
        )
        out["connected"] = True
    except Exception as exc:
        report.append(f"cannot connect to soffice on port {port}: {exc}")
        _write(result_path, out)
        return 1

    smgr = ctx.ServiceManager

    out["vba_config"] = read_vba_config(ctx, report)

    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    props = []
    for name, value in (
        ("MacroExecutionMode", MACRO_EXEC_ALWAYS_NO_WARN),
        ("Hidden", True),
        ("UpdateDocMode", 0),
    ):
        p = PropertyValue()
        p.Name = name
        p.Value = value
        props.append(p)

    report.append(f"loading {workbook} (MacroExecutionMode={MACRO_EXEC_ALWAYS_NO_WARN})")

    doc = None
    try:
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(workbook), "_blank", 0, tuple(props)
        )
    except Exception as exc:
        report.append(f"load failed: {type(exc).__name__} {exc}")

    if doc is None:
        report.append("document did not load")
        _write(result_path, out)
        return 1

    report.append("document loaded")

    try:
        out["libraries"] = describe_libraries(doc, report)
        install_bridge(doc, report)

        factory = smgr.createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", ctx
        )
        msp = factory.createScriptProvider(doc)

        for group, module, func in TESTS:
            status, detail = invoke_wrapper(msp, func)
            out["results"].append(
                {
                    "group": group,
                    "module": module,
                    "function": func,
                    "status": status,
                    "detail": detail,
                }
            )
    except Exception:
        report.append("probe crashed:\n" + traceback.format_exc()[:1500])
    finally:
        try:
            doc.close(True)
            report.append("document closed")
        except Exception:
            pass

    _write(result_path, out)
    return 0


def _write(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


if __name__ == "__main__":
    sys.exit(main())
