
root_path = "E:\HOC\VBS\VBS_Useful_Code\New folder"

ext_name = ".txt"

Sub recursiveDeleteMatchedExtensionFile(root)
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set sf = root.SubFolders

    If sf.Count > 0 Then
        For Each f1 in sf
            Call recursiveDeleteMatchedExtensionFile(f1)
        Next
    End If

    Set file_list = root.Files
    If file_list.Count > 0 Then
        For Each found_file in file_list
            If ("." & fs.GetExtensionName(found_file)) = ext_name Then
                Call fs.DeleteFile(found_file)
            End If
        Next
    End If
End sub



Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(root_path)

Call recursiveDeleteMatchedExtensionFile(f)

Set fs = Nothing
Set f = Nothing

MsgBox("Done")