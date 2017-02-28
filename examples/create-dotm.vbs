Option Explicit

Const wdFormatXMLTemplateMacroEnabled = 15

Dim appWd
Dim doc

Dim docPath
Dim arrVBA
Dim refPaths
Dim prjFiles
Dim prjRoot
Dim iX

prjRoot = "C:\Users\jgantner\Documents\VBAStuff\vba\"
prjFiles = Array( _
            "Macros.bas", _
            "RibbonControl.bas", _
            "Main.frm" _
            )

refPaths = Array("C:\Program Files\Common Files\System\ado\msado28.tlb")

docPath = "C:\Users\anon\Documents\VBAStuff\blah.dotm"

Set appWd = CreateObject("Word.Application")
'appWd.Visible = True

Set doc = appWd.Documents.Add()

For iX = LBound(prjFiles) To UBound(prjFiles)
    doc.VBProject.VBComponents.Import prjRoot & prjFiles(iX)
Next 'iX

For iX = LBound(refPaths) To UBound(refPaths)
    doc.VBProject.References.AddFromFile refPaths(iX)
Next 'iX

doc.SaveAs docPath, wdFormatXMLTemplateMacroEnabled

doc.Close

Set doc = Nothing

appWd.Quit
Set appWd = Nothing


Dim zipPath
zipPath = """C:\Program Files\7-Zip\7z.exe"" l " & docPath

Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
iX = objShell.Run (zipPath, 1, True)

