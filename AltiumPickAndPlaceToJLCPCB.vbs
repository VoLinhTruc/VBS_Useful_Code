
' file_in_path = "E:\Full-product\MQSC_1.0\Hardware\MQSC_1.0.H1\Project Outputs for MQSC_1.0.H1\Pick Place for PCB2.csv"
' file_out_path = "E:\Full-product\MQSC_1.0\Hardware\MQSC_1.0.H1\Project Outputs for MQSC_1.0.H1\Pick Place for PCB2 Final.csv"

file_in_path = "E:\Full-product\MQSC_1.0\Hardware\MQSC_1.0.H_Lite\Project Outputs for MQSC_1.0.H_Lite\Pick Place for PCB2.csv"
file_out_path = "E:\Full-product\MQSC_1.0\Hardware\MQSC_1.0.H_Lite\Project Outputs for MQSC_1.0.H_Lite\Pick Place for PCB2 Final.csv"

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(file_in_path,1)
strFileText = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

start_index = InStr(strFileText,"Designator")

result = Mid(strFileText,start_index)
result = Replace(result,"""","")
result = Replace(result,"Center-","Mid ")
result = Replace(result,"(mm)","")
result = Replace(result,"TopLayer","Top")
result = Replace(result,"BottomLayer","Bottom")


Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(file_out_path,2,true)
objFileToWrite.WriteLine(result)
objFileToWrite.Close
Set objFileToWrite = Nothing