Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = "\\sac-filsrv1\Projects\Structural-028\Projects\LEB\9.0_Const_Svcs\Mortenson\Submittals\Submittal - 270 - LEBADMIN_061543_Calc_Area AB_Admin CLT Roof Structural Calculations & Gravity Beam Analysis"

If fso.FolderExists(folderPath) Then
    WScript.Echo "Folder EXISTS"
    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files
        WScript.Echo file.Name
    Next
Else
    WScript.Echo "Folder NOT FOUND: " & folderPath
End If
