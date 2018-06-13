' Usage:
'  WScript.exe Compose.vbs <file> <path>

' Converts all modules, classes, forms, queries and macros in a directory created by "Decompose.vbs"
' and composes then into an Access file. This overwrites any existing Modules with the
' same names without warning!!!
' Requires Microsoft Access.

Option Explicit

const acQuery = 1
const acForm = 2
const acModule = 5
const acMacro = 4
const acReport = 3

Const acCmdCompileAndSaveAllModules = &H7E

' BEGIN CODE
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

dim sAccessFilename
If (WScript.Arguments.Count = 0) then
    MsgBox "Please specify the filename!", vbExclamation, "Error"
    Wscript.Quit()
End if
sAccessFilename = fso.GetAbsolutePathName(WScript.Arguments(0))

Dim sPath
If (WScript.Arguments.Count = 1) then
    sPath = ""
else
    sPath = WScript.Arguments(1)
End If


importModulesTxt sAccessFilename, sPath

If (Err <> 0) and (Err.Description <> NULL) Then
    MsgBox Err.Description, vbExclamation, "Error"
    Err.Clear
End If

Function importModulesTxt(sAccessFilename, sImportpath)
    Dim myComponent
    Dim sModuleType
    Dim sTempname
    Dim sOutstring

    ' Build file and pathnames
    dim myType, myName, myPath, sSubAccessFilename
    myType = fso.GetExtensionName(sAccessFilename)
    myName = fso.GetBaseName(sAccessFilename)
    myPath = fso.GetParentFolderName(sAccessFilename)

    ' if no path was given as argument, use a relative directory
    If (sImportpath = "") then
        sImportpath = myPath & "\Source\"
    End If
    sSubAccessFilename = sImportpath & myName & "_sub." & myType

    ' check for existing file and ask to overwrite with the stub
    if (fso.FileExists(sAccessFilename)) Then
        WScript.StdOut.Write sAccessFilename & " exists. Overwrite? (y/n) "
        dim sInput
        sInput = WScript.StdIn.Read(1)
        if (sInput <> "y") Then
            WScript.Quit
        end if

        fso.CopyFile sAccessFilename, sAccessFilename & ".bak"
    end if

    fso.CopyFile sSubAccessFilename, sAccessFilename

    ' launch MSAccess
    WScript.Echo "Starting Access..."
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")
    WScript.Echo "Opening " & sAccessFilename & " ..."
    If (Right(sSubAccessFilename,4) = ".adp") Then
        oApplication.OpenAccessProject sAccessFilename
    Else
        oApplication.OpenCurrentDatabase sAccessFilename
    End If
    oApplication.Visible = false

    Dim folder
    Set folder = fso.GetFolder(sImportpath)

    ' load each file from the import path into the stub
    Dim myFile, objectname, objecttype
    for each myFile in folder.Files
        objecttype = fso.GetExtensionName(myFile.Name)
        objectname = fso.GetBaseName(myFile.Name)
        WScript.Echo "RECOMPOSING " & objectname & " (" & objecttype & ")"

        if (objecttype = "form") then
            oApplication.LoadFromText acForm, objectname, myFile.Path
        elseif (objecttype = "bas") then
            oApplication.LoadFromText acModule, objectname, myFile.Path
        elseif (objecttype = "mac") then
            oApplication.LoadFromText acMacro, objectname, myFile.Path
        elseif (objecttype = "report") then
            oApplication.LoadFromText acReport, objectname, myFile.Path
	elseif (objecttype = "query") then
	    oApplication.LoadFromText acQuery, objectname, myFile.Path
        end if

    next

    oApplication.RunCommand acCmdCompileAndSaveAllModules
    oApplication.Quit

    MsgBox ("Compose completed!")

End Function

Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.source & ":" & vbCrLf & _
               "    Description: " & Err.Description & vbCrLf & _
               "    Code: " & Err.Number & vbCrLf
    getErr = strError
End Function