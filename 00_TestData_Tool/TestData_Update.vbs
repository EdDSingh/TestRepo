Set ObjExcel = CreateObject("Excel.Application")

ObjExcel.Visible = True
ObjExcel.DisplayAlerts = False

Set shell = CreateObject("WScript.Shell")
shell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

Set ObjExcelFile = ObjExcel.Workbooks.Open(scriptdir &"\"& "Test_1.xlsx")
Set ObjExcelSheet = ObjExcelFile.Sheets("Sheet1")
Set shell = Nothing
						
Dim TimeDateDetails

Dim i,loop_end,column_indexA,range,FileName,SheetName,olddata,newData,TCstatus
i=2

ConvertTemp()
For i = 2 To 10000
	ObjExcelFile.Sheets("Sheet1").Activate
	TCstatus = ObjExcelSheet.Range("B"&i).Value
	FileName = ObjExcelSheet.Range("C"&i).Value & ObjExcelSheet.Range("D"&i).Value

	If FileName = "" Then
		Exit For
	Else
			IF TCstatus="Y" OR TCstatus="y" Then

					FileType = ObjExcelSheet.Range("E"&i).Value
					Dim TempNewOrder 
					TempNewOrder = ObjExcelSheet.Range("G"&i).Value
					ObjExcelSheet.Range("F"&i).Value = TempNewOrder
					olddata = ObjExcelSheet.Range("F"&i).Value
					
					ObjExcelSheet.Range("G"&i).Value = ObjExcelSheet.Range("A"&i).Value & TimeDateDetails
					newData = ObjExcelSheet.Range("G"&i).Value
					ObjExcelFile.Save
				If FileType = "Excel" Then
					Set ObjChangeFile = ObjExcel.Workbooks.Open(FileName)
					For x=1 To objExcel.Worksheets.count
							Set ObjChangeSheet = ObjChangeFile.Sheets(x)
							ObjChangeFile.Sheets(x).Activate
							Set objRange = ObjChangeSheet.UsedRange
							objRange.Replace olddata, newData
					Next
					ObjChangeFile.Save
					ObjChangeFile.Close
				ElseIf FileType = "CSV" Then
						Set objFSO = CreateObject("Scripting.FileSystemObject")
						Set objFile = objFSO.OpenTextFile(FileName, 1)
						strText = objFile.ReadAll
						objFile.Close
						strNewText = Replace(strText, olddata, newData)
						Set objFile = objFSO.OpenTextFile(FileName, 2)
						objFile.WriteLine strNewText 
						objFile.Close
						Set shell = CreateObject("WScript.Shell")
						shell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
						shell.Run "AutoIT_RemoveBlankLine.bat "&FileName
						WScript.Sleep 1000
						Set objFSO = Nothing
						Set objFile = Nothing
						Set shell = Nothing
				Else
					MsgBox("Please Enter Correct File Type (Excel or CSV)")
				End If
			End If
	End IF
Next


ObjExcelFile.Close

Sub ConvertTemp()
		Dim strNow, strDD, strMM, strYYYY, strFulldate
		strYYYY = DatePart("yyyy",Now())
		strMM = Right("0" & DatePart("m",Now()),2)
		strDD = Right("0" & DatePart("d",Now()),2)
		fulldate = strYYYY & strMM & strDD

		Dim hours 
		hours = Right("0" & DatePart("h",Now()),2)
		Dim minutes
		minutes = Right("0" & DatePart("n",Now()),2)
		Dim seconds
		seconds = Right("0" & DatePart("s",Now()),2)
		
		Dim exacttime
		exacttime = hours & minutes & seconds
		
		TimeDateDetails = "D" & fulldate & "T" & exacttime
End Sub