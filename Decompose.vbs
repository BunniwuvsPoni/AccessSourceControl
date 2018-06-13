' Usage:
'  CScript.exe Decompose.vbs <input file> <path>

' Converts all modules, classes, forms, queries and macros from an Access file <input file> to
' text and saves the results in separate files to <path>.  Requires Microsoft Access.
'

Option Explicit

const acQuery = 1
const acForm = 2
const acModule = 5
const acMacro = 4
const acReport = 3

' BEGIN CODE
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

dim sAccessFilename
If (WScript.Arguments.Count = 0) then
    MsgBox "Please specify the filename!", vbExclamation, "Error"
    Wscript.Quit()
End if
sAccessFilename = fso.GetAbsolutePathName(WScript.Arguments(0))

Dim sExportpath
If (WScript.Arguments.Count = 1) then
    sExportpath = ""
else
    sExportpath = WScript.Arguments(1)
End If


exportModulesTxt sAccessFilename, sExportpath

If (Err <> 0) and (Err.Description <> NULL) Then
    MsgBox Err.Description, vbExclamation, "Error"
    Err.Clear
End If

Function exportModulesTxt(sAccessFilename, sExportpath)
    Dim myComponent
    Dim sModuleType
    Dim sTempname
    Dim sOutstring

    dim myType, myName, myPath, sSubAccessFilename 
    myType = fso.GetExtensionName(sAccessFilename)
    myName = fso.GetBaseName(sAccessFilename)
    myPath = fso.GetParentFolderName(sAccessFilename)

    If (sExportpath = "") then
        sExportpath = myPath & "\Source\"
    End If
    sSubAccessFilename  = sExportpath & myName & "_sub." & myType

    WScript.Echo "Deleting existing folder..."
    Dim exists
    exists = fso.FolderExists(sExportpath)

    If (exists) then
	If Right(sExportpath, 1) = "\" then
	   Dim truncate_one
	   truncate_one = Left(sExportpath, Len(sExportpath) - 1)
	   fso.DeleteFolder truncate_one
	Else
	   fso.DeleteFolder sExportpath
	End If
    End If

    WScript.Echo "Copy stub to " & sSubAccessFilename  & "..."
    On Error Resume Next
        fso.CreateFolder(sExportpath)
    On Error Goto 0
    fso.CopyFile sAccessFilename, sSubAccessFilename 

    WScript.Echo "Starting Access..."
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")
    WScript.Echo "Opening " & sSubAccessFilename  & " ..."
    oApplication.OpenCurrentDatabase sSubAccessFilename 

    oApplication.Visible = false

    dim dctDelete
    Set dctDelete = CreateObject("Scripting.Dictionary")
    WScript.Echo "Exporting..."
    Dim myObj
    For Each myObj In oApplication.CurrentDb.QueryDefs
        Wscript.Echo "Exporting QUERY " & myObj.Name
        oApplication.SaveAsText acQuery, myObj.Name, sExportpath & "\" & myObj.Name & ".query"
	dctDelete.Add "QU" & myObj.Name, acQuery
    Next
    For Each myObj In oApplication.CurrentProject.AllForms
        WScript.Echo "Exporting FORM " & myObj.fullname
        oApplication.SaveAsText acForm, myObj.fullname, sExportpath & "\" & myObj.fullname & ".form"
        oApplication.DoCmd.Close acForm, myObj.fullname
        dctDelete.Add "FO" & myObj.fullname, acForm
    Next
    For Each myObj In oApplication.CurrentProject.AllModules
        WScript.Echo "Exporting MODULE " & myObj.fullname
        oApplication.SaveAsText acModule, myObj.fullname, sExportpath & "\" & myObj.fullname & ".bas"
        dctDelete.Add "MO" & myObj.fullname, acModule
    Next
    For Each myObj In oApplication.CurrentProject.AllMacros
        WScript.Echo "Exporting MACRO " & myObj.fullname
        oApplication.SaveAsText acMacro, myObj.fullname, sExportpath & "\" & myObj.fullname & ".mac"
        dctDelete.Add "MA" & myObj.fullname, acMacro
    Next
    For Each myObj In oApplication.CurrentProject.AllReports
        WScript.Echo "Exporting REPORT " & myObj.fullname
        oApplication.SaveAsText acReport, myObj.fullname, sExportpath & "\" & myObj.fullname & ".report"
        dctDelete.Add "RE" & myObj.fullname, acReport
    Next

    WScript.Echo "Deleting..."
    dim sObjectname
    For Each sObjectname In dctDelete
        WScript.Echo "OBJECT " & Mid(sObjectname, 3)
        oApplication.DoCmd.DeleteObject dctDelete(sObjectname), Mid(sObjectname, 3)
    Next

    oApplication.CloseCurrentDatabase
    oApplication.CompactRepair sSubAccessFilename , sSubAccessFilename  & "_"
    oApplication.Quit

    fso.CopyFile sSubAccessFilename  & "_", sSubAccessFilename 
    fso.DeleteFile sSubAccessFilename  & "_"

    MsgBox ("Decompose completed!")

End Function

Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.source & ":" & vbCrLf & _
               "    Description: " & Err.Description & vbCrLf & _
               "    Code: " & Err.Number & vbCrLf
    getErr = strError
End Function