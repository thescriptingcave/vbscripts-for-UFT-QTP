'This utility is designed to convert multiple columns excel datasheet in to one singe column excel datatable with appropriate titles
'It put  it in the appropriate format to KDSE

'@@@@@@@@@@@@@Aurthor : Demeke Bogale @@@@@@@@@@@@@@@@@

'strPath1 = "C:\DataTables\datatable2.xls"

Const 	xlContinuous = 1
strPath2 = "C:\DataTables\datatable.xls"
DataTableName =  Environment.Value("DataTableName")

Function DeleteExcelFile
  	file_location = "C:\DataTables\datatable.xls"
	Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile(file_location)

End Function

Function CreateFolderDatatable
	Set f = fso.CreateFolder("C:\DataTables\")
	CreateFolderDatatable  = f.Path
End Function

Function CreateNewExcelFile
        Set xlBook  =  oExcel.workbooks.add
		Set xlSheet = xlBook.activesheet
            	xlBook.saveas "C:\DataTables\datatable.xls",56
				oExcel.quit
End Function

	Set fso  =  CreateObject("Scripting.FileSystemObject" )
     Set oExcel = createobject("Excel.Application")
	 'Set objFolder = fso.GetFolder("C:\DataTables\")
			If Not fso.FolderExists("C:\DataTables\") then
					CreateFolderDatatable
		end if 'Else
            	If fso.FileExists( "C:\DataTables\datatable.xls") Then
						DeleteExcelFile
						CreateNewExcelFile
				Else
						 CreateNewExcelFile
			    End if
			'End If

Set oSource = oExcel.Workbooks.Open(DataTableName)
Set oDest  = oExcel.Workbooks.Open(strPath2)
set oD = oDest.Worksheets(1)

WorksheetsCount = oSource.Worksheets.count


Dim numberOfTables
'numberOfTables = 0
For t = 1 to WorksheetsCount
			set  activesheet = oSource.Worksheets(t)		'make the t - th sheet active
			If  Ucase(left(Activesheet.name, 5)) = "TABLE" then
					numberOfTables = numberOfTables + 1
			End if
Next

sheetsNeeded = numberOfTables + 1 				'1 is added for the first sheet where the data table is going to be created  from all the sheets (excluding the tables)

'Add Three Sheets in the destination sheet where you have sheets refered by oD
If SheetsNeeded > 3 Then
	sheetsToAdd = sheetsNeeded - 3
    oExcel.activeworkbook.sheets.add,oDest.Sheets(oDest.Sheets.count),sheetsToAdd 		''Add as maney sheets as specified by  SheetsNeeded at the end of the
																																													'the existing sheets
end if


'oD.Columns("D:D").NumberFormat="@"
pos = 1
oD.Cells(pos, 1).Value = "Browser"
oD.Cells(pos, 2).Value = "Page"
oD.Cells(pos, 3).Value = "Object"
oD.Cells(pos, 4).Value = "Action"
oD.Cells(pos, 5).Value = "Data"
oD.Cells(pos, 6).Value = "Step_Name"
oD.Cells(pos, 7).Value = "Comment"			'

For m = 1 to 7
	oD.Cells(pos, m).interior.ColorIndex = 17
Next

row = 5
i = 1
tableSheetPos = 2
oD.Range("E1:IV65536").numberformat = "@"
flag1 = false
flag2 = true

Do
    set  activesheet = oSource.Worksheets(i)
	set  cells = activesheet.usedrange
	usedColumnCount = activesheet.usedrange.columns.count
	usedRowsCount = activesheet.usedrange.rows.count

   	If (left(activesheet.Name,5) =  "Table" ) then
			if(tableSheetPos <= sheetsNeeded) then
	           if(flag2) then	'Once it copied the tables, the next time it gets them it should not do anything
		       flag1 = true

			   set oD = oDest.Worksheets(tableSheetPos)
			   oD.name = activesheet.name		'Rename the destination sheet

			  For tRow = 1 to usedRowsCount
			    For tCol = 1 to usedColumnCount
                    oD.Cells(tRow, tCol).Value = Cells(tRow, tCol).Value
                Next
	         Next
			set oD = oDest.Worksheets(1)
		 end if
	end if


		else
For col = 1 to usedColumnCount
	Data =  rtrim(ucase(Cells(row, col).Value ))
    DataAsIs =  rtrim(Cells(row, col).Value )

	If  Data  <> "" Then
	   sBrowser = Cells(1, col).Value
		sPage = Cells(2, col).Value
		sObject = Cells(3, col).Value
		action = Cells(4, col).Value

		pos = pos + 1
		oD.Cells(pos, 1).Value = sBrowser
		oD.Cells(pos, 2).Value = sPage
		oD.Cells(pos, 3).Value = sObject							'Ucase(sObject)
		'oD.Cells(pos, 3).Value = sObject
		oD.Cells(pos, 4).Value = ucase(action)			'3 is changed to 4

		 If Ucase(Data) <> "SUBMIT" Then
			If  Ucase(Data) <> "GETTABLE" Then
		  	   oD.Cells(pos, 5).Value = DataAsIs					'Data
			else
			  oD.Cells(pos, 4).Value = DataAsIs					' Data
			end if
	else
			oD.Cells(pos, 4).Value = DataAsIs									' Data
						 End If
                    End If
			Next
	End if

	If i >= WorksheetsCount Then
			i = 1
	else
			 i = i+1
	End If
	If i = 1 Then
		 row = row + 1
	End If
	If tableSheetPos >  SheetsNeeded Then		'(Why????When it gets any table after this it just skips to the next sheet without doing anything)
		flag2 = false
	End If
	If flag1= true Then
			tableSheetPos = tableSheetPos +1
			flag1 = false	'tableSheetPos should be incremented only when control is entered in the top first if loop. Thus-  flag1 must be true
									' only when control is entered in the top first if loop
        End If

Loop	Until ((row > usedRowsCount))

'Auto fit sheet1 of the destination workbook
set oD = oDest.Worksheets(1)
Set objRange = oD.Range("A1:" &"F" & pos)
objRange.Font.Name = "System"
oD.Columns("A:A").Autofit()
oD.Columns("B:B").Autofit()
oD.Columns("C:C").Autofit()
oD.Columns("D:D").Autofit()
oD.Columns("E:E").Autofit()

objRange.Borders.LineStyle = xlContinuous


oDest.Save
oDest.Close
oSource.Close
oExcel.Quit
