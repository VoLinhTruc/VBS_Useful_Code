' This program delete all the folder which has the name matched in the folder_name_list below
' Use "Call addItemToArray(folder_name_list, <name_of_the_folder_you_want_to_delete>)" to add the folder name you want to delete

Redim folder_name_list(0)
Call addItemToArray(folder_name_list, "Debug")
Call addItemToArray(folder_name_list, ".vs")


Function addItemToArray(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
    addItemToArray = arr
End Function

Set WshShell = WScript.CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(WshShell.CurrentDirectory)

Sub recursiveFolderDelete(root)
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set sf = root.SubFolders

    If sf.Count > 0 Then
        For Each f1 in sf
            For Each folder_name in folder_name_list
                If f1.name = folder_name Then
                    fs.DeleteFolder(f1)
                ELse
                    Call recursiveFolderDelete(f1)
                End If
            Next
        Next
    End If
End sub

Call recursiveFolderDelete(f)

Set fs = Nothing
Set f = Nothing

MsgBox("Done")