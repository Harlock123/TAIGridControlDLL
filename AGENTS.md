# AGENTS.md

Notes for anyone, human or AI agent, changing this code. `README.md` has the back story and the version history,
and `AncientHistory.txt` the change log from the original VB.NET years; this file covers how the code is laid out
and how to build and change it safely.

## What this is

`TAIGridControl`, a WinForms grid control (a `UserControl`) for programmatic, VB6-style use: fill it cell by cell
or from a SQL query, sort, format, print, and export to Excel, HTML or SQL scripts. It is C# on .NET Framework 4.8,
built as the library `TAIGridControl2.dll`, with a small WinForms test program. It builds and runs on Windows only.

## Commands

Build on Windows with Visual Studio, or from a Developer Command Prompt:

```bat
nuget restore TAIGridControlDLL.sln
msbuild TAIGridControlDLL.sln /p:Configuration=Release
```

- The library builds to `bin\` (both configurations). The NuGet packages (ClosedXML.Signed 0.95.4,
  DocumentFormat.OpenXml 2.11.0, ExcelNumberFormat 1.0.10) come from `packages.config` into `packages\`.
- `TAIGridControlDLL.csproj` also lists COM references (`Microsoft.Office.Core`, `VBIDE`), so the build machine
  needs the matching Office interop libraries registered.
- There are no automated tests. `GridTests` is a WinForms program (`frmMain`) with buttons that exercise the
  grid by hand; it is the solution's startup project.

## Layout

- `TAIGridControl.cs`: the whole control, about 21,000 lines: data, drawing, mouse and keyboard handling,
  sorting, context menus, printing and the public API. `TAIGridControlParameter` is defined at its end.
- `frm*.cs` (with `.resx`): dialogs the grid opens: column tear-away windows, frequency distribution, grid
  properties, page setup, and export to HTML, SQL and Excel (`frmExcelOutput`, which uses ClosedXML).
- `*.vb`: the original VB.NET sources, kept for reference. The C# project does not compile them; change the
  `.cs` files.
- `My Project/MyNamespace.Static.Designer.cs`: C# stand-ins for the VB `My` namespace, left from the conversion.
- `TAIGridControl2.xml`: the XML documentation file the build writes; it is committed.
- `Documentation/`: an old Sandcastle Help File Builder project and its built `TAIGridControl2DLL.chm`. The README
  notes this project no longer builds with current Visual Studio.
- `GridTests/`: the manual test program.

## Rules that keep it working

- **Keep the public API stable.** Other WinForms applications use this control; rename or remove public methods,
  properties and events only on purpose, and add overloads rather than change existing signatures.
- **Excel export goes through ClosedXML**, not Excel COM automation. Keep it that way so exporting does not need
  Excel installed.
- **New source files must be added to the `.csproj` by hand.** It is an old-style project with explicit
  `Compile` and `EmbeddedResource` items.
- **The version** is in `AssemblyInfo.cs` (now 4.0.0.0). The README's version history skips from 2.0.14.0 to
  4.0.0.0; the versions in between are not listed.

## Making a change

- Build the solution with no new warnings, then run `GridTests` and exercise the change by hand.
- Add a line to the version history in `README.md` for user-visible changes.

There is no CI workflow in this repository.
