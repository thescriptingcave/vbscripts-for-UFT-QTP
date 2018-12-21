'This Utility updates the datasheet with values found after a test executed
'*****************************************************************************************************************************
Dim FoundItemsNotBlank, AllFoundItems, Loc
Dim  AllValues,  ValuesLocation, SoureDestMatch
Dim z, cnt
cnt = 0
'*******************************************************************************
Function CreateFolderDatatable
	Set f = fso.CreateFolder("C:\DataSheetsUpdated\")
	CreateFolderDatatable  = f.Path
End Function
'*************************************************************************************************************************

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Function CreateMappingTable(ByRef oS, usedRowCountoS, AllValues,ValuesLocation, FoundItemsNotBlank, AllFoundItems)
For Loc  = 2 to usedRowCountoS
	 ValueFound = oS.Cells(Loc, 7).Value
					'print ValueFound 		'Debug
	 AllValues.Add Loc, ValueFound 'Put all values in the DataFound column (7th column) and the location of each value in the datatable

	 AllFoundItems  = AllValues.Items ' All Values (empty and not empty) found in the DataFound (7th) column
     ' print dict(k-1)
	If  ValueFound <> ""      Then
		ValuesLocation.Add Loc, ValueFound 'Not empty values and location in the datatable
		FoundItemsNotBlank = ValuesLocation.items ' All Values not empty, found in the DataFound (7th) column
   end if
		Next
End Function
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Const 	xlContinuous = 1
' lastVal = "SIGNOFF"			'original

DataTableName =  Environment.Value("DataTableName")
DataSheetName  =  Environment.Value("DataTableName2")

Set oExcel = createobject("Excel.Application")
'oExcel.Visible = true debug
Set AllValues   =  CreateObject("Scripting.Dictionary")
Set ValuesLocation  = CreateObject("Scripting.Dictionary")

Set fso  =  CreateObject("Scripting.FileSystemObject" )
Set oSource = oExcel.Workbooks.Open(DataTableName)
Set oDest  = oExcel.Workbooks.Open(DataSheetName)

set oS   = oSource.Worksheets(1)
usedRowCountoS = oS.usedrange.rows.count 'Total number of rows in the source/delta datatable
'newly added statement
lastValue = oS.cells(usedRowCountoS, 1)
If usedRowCountoS = 65536  Then			'changed 01/05/2012 due the problem with excel 2010
		usedRowCountoS = oS.UsedRange.Find(lastVal).Row			'What ever value we give to lastVal, this will always return the last row number
												'included to avoid problems with rows below the last row. Makes the spreadsheet as if it has only
												'usedRowCountoS rows than 65536 rows.
 End If
'print "used Rows:="& usedRowCountoS

If Not fso.FolderExists("C:\DataSheetsUpdated\") then 'Create Datatable folder if it is not already created
		CreateFolderDatatable
End if

CreateMappingTable oS, usedRowCountoS, AllValues,ValuesLocation, FoundItemsNotBlank, AllFoundItems 'Call Mapping Table

Counter = 1
WorksheetsCount =   oDest.Worksheets.count
					'print "Sheet Count: ="&WorksheetsCount

'********new code************************

For z = 1 to WorksheetsCount
	Set currentSheet = oDest.Sheets(z)
	currentSheetName = currentSheet.Name
					'print currentSheetName
	leftFiveCharacters = Left(currentSheetName, 5)
					'print leftFiveCharacters
	If  ucase(leftFiveCharacters) <>	ucase("Table")Then
			cnt = cnt + 1
	End If
Next
WorksheetsCount = cnt

'********End of new code************************

'row = 4		 	'for mainframe actual data values, in the datasheet, start from row 4. The first 3 rows
							'provide header information
row = 5			'For the web app, actual data values, in the datasheet, start from row 5.
					'The first 3 rows provide header information
i = 1
		Do
	set  activesheet = oDest.Worksheets(i)
	set  cells = activesheet.usedrange
	usedColumnCount = activesheet.usedrange.columns.count
	usedRowsCount = activesheet.usedrange.rows.count

For col = 1 to usedColumnCount
	Data =  rtrim(ucase(Cells(row, col).Value ))

	If  Data  <> "" Then 'Ignore empty cells/values
	'	Cells(row,col).Interior.ColorIndex = i+2
		If  ValuesLocation.Exists(Counter +1) Then 'Find the location  of the data to be updated in the datasheet
			Cells(row, col).Interior.ColorIndex = 4 'change the color of the cell to be updated
			Cells(row, col).NumberFormat= "@"
			Cells(row, col).Value  = AllFoundItems(counter-1) 'Get corrosponding value for that location
		 End If
		 Counter  =  Counter +1
	'	pos = pos + 1
	end if
 Next
	 'print "Value of counter: " & counter


 If i >= WorksheetsCount Then 'go until last sheet for each row
	i = 1
		else 'increament sheet number until it reaches the last sheet
	 i = i+1
End If

If i = 1 Then 'increament row number after the last sheet
		 row = row + 1
End If
Loop Until row > usedRowsCount 'loop until all rows are executed

SoureDestMatch  = True
'print "Counter: =  " & Counter
'print "Datasheet Counter:= " & AllValues.Count +1
	If Counter <>  AllValues.Count +1 Then
	SoureDestMatch  =  False
	Reporter.ReportEvent micFail, "Total number of Rows",  "Total number of Rows in the Source datatable don't much with No. of non empty cells in the Datasheet to be Updated"
	'Msgbox "Total number of Rows in the Source datatable don't much with non emty cells in Datasheet to be Updated"
	'Exit test
	End If

	If  SoureDestMatch Then
	oDest.Save
	'oSource.Save
	Else
	oDest.Saved 	= 	True
	End if

oDest.Close    
oSource.Close
oExcel.Quit
