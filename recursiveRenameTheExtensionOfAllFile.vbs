
root_path = "C:\C_Cpp_LIB\Bico_MQTT\src\New folder (2)"

src_ext_name = ".cpp"
dsn_ext_name = ".c"

Sub recursiveFileRenameExtension(root)
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set sf = root.SubFolders

    If sf.Count > 0 Then
        For Each f1 in sf
            Call recursiveFileRenameExtension(f1)
        Next
    End If

    Set file_list = root.Files
    If file_list.Count > 0 Then
        For Each found_file in file_list
            If ("." & fs.GetExtensionName(found_file)) = src_ext_name Then
                found_file.Name = Replace(found_file.Name, src_ext_name, dsn_ext_name)
            End If
        Next
    End If
End sub



Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(root_path)

Call recursiveFileRenameExtension(f)

Set fs = Nothing
Set f = Nothing

MsgBox("Done")