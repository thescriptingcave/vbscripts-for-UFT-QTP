'AgileSearchTool Developed by Robert Harding, Robert Harding Consulting.  2/24/2011



'********************* Constants corresponding to Sheet 1 of the control file***************
Const xlCellTypeVisible =12
Const BrowserColumn = 1
Const PageColumn = 2
Const ObjectColumn = 3
Const ActionColumn = 4
Const DataColumn = 5
Const CommentColumn = 6
'***************************************************************************************

strB = "C:\Temp\controlfile.xlsx"

'*************************** Quality Center Test Folder Path ************************************************
'     qctestfolderpath ="Subject\System Test Scripts\UI\Automation" '\5.Post Automation/Batch Verification\1.Post Automation Verification"
'qctestfolderpath ="Subject\Rich Cleanup"
qctestfolderpath ="Subject\QTP DEMO Scripts"
'***************************************************************************************************



Set excel_CF = createobject("Excel.Application")
Set excelwkbk_CF = excel_CF.Workbooks.Open(strB)
Set excelsheets_CF =excelwkbk_CF.Worksheets
excelsheets_CF(2).Columns("A:Z").NumberFormat ="@"


excel_CF.Visible =True


rowC = 0

Set tdc = qcutil.QCConnection
Set TreeMgr = tdc.TreeManager
Set SubjRoot = TreeMgr.NodeByPath(qctestfolderpath)
Set TestFact = tdc.TestFactory

Set SubjectNodeList = SubjRoot.FindChildren("", False, "")

For Each oSubjectNode In SubjectNodeList
    'print oSubjectNode.name

Set TestList =GetTestObjects(oSubjectNode.path)


      For each tst in TestList


counter = 0
sflag =False

Set aTestObject = tdc.TestFactory.item(tst.id)
Set AttachmentFact = aTestObject.Attachments
Set AttachmentColl = AttachmentFact.NewList("")

For each mfile in AttachmentColl
counter = counter +1
mysplit = split(mfile.name,"_",-1)
myval = lcase(ubound(mysplit))

    If myval = 2 and  (lcase(mysplit (2)))="datatable.xls" Then

    Set theAttach = AttachmentColl.Item(counter)
        theAttach.Load True, ""
        'print theAttach.id
        theFileName = theAttach.FileName
        theFileSize = theAttach.FileSize
        'print theFileSize
        If theAttach.FileSize = 0 Then
            Exit for
        End If
        thePath = Left(theFileName, InstrRev(theFileName, "\")-1)
        sflag = True
        Exit for
    End If

Next

If sflag = True Then



theAttach.Load True, ""
theFileName = theAttach.FileName
thePath = Left(theFileName, InstrRev(theFileName, "\")-1)
Set excel = createobject("Excel.Application")
Set excelwkbk = excel.Workbooks.Open(theFileName)
Set excelsheets = excelwkbk.Sheets
excel.Visible =False
excelsheets(1).select
row_count =excelsheets(1).UsedRange.rows.count
column_count =excelsheets(1).UsedRange.columns.count
FileControlRowTotal =excelsheets_CF(1).UsedRange.Rows.Count

For FileControl = 2 to FileControlRowTotal
    aBrowser = trim(excelsheets_CF(1).Cells(FileControl,1))
    aPage = trim(excelsheets_CF(1).Cells(FileControl,2))
    aObject = trim(excelsheets_CF(1).Cells(FileControl,3))
    Action = trim(excelsheets_CF(1).Cells(FileControl,4))
    Data= excelsheets_CF(1).cells(FileControl,5)
    Comment = excelsheets_CF(1).cells(FileControl,6)
        If excel.ActiveSheet.AutofilterMode = True  Then
    	  excel.ActiveSheet.AutofilterMode = False
          End If

'Browser only
If aBrowser <> "" and aPage =  "" and aObject = "" and Action = "" and Data = "" Then
Set FilterRange =excelsheets(1).UsedRange

FilterRange.Autofilter 1, aBrowser

RowCount = filterRange.SpecialCells(xlCellTypeVisible).Rows.Count
For each mcell in FilterRange.Specialcells(xlCellTypeVisible).Rows
	If mcell.Row > 1 Then
		  x =GetEmptyRowNumber(excelsheets_CF(2))
		  excelsheets_CF(2).cells(x+1,1) = tst.name
		  excelsheets_CF(2).cells(x+1,2) = tst.id
		  excelsheets_CF(2).cells(x+1,3) = excel.Cells(mcell.Row, BrowserColumn).Value
		  excelsheets_CF(2).cells(x+1,9) = mcell.Row
	End If

next

End If

'Page Only
If aBrowser = "" and aPage <>  "" and aObject = "" and Action = "" and Data = "" Then

Set FilterRange =excelsheets(1).UsedRange

	FilterRange.Autofilter 2, aPage

RowCount = filterRange.SpecialCells(xlCellTypeVisible).Rows.Count

For each mcell in FilterRange.Specialcells(xlCellTypeVisible).Rows

	If mcell.Row > 1 Then

		x =GetEmptyRowNumber(excelsheets_CF(2))
		excelsheets_CF(2).cells(x+1,1) = tst.name
		excelsheets_CF(2).cells(x+1,2) = tst.id
		excelsheets_CF(2).cells(x+1,3) = excel.Cells(mcell.Row, BrowserColumn).Value
		excelsheets_CF(2).cells(x+1,4) = excel.Cells(mcell.Row, PageColumn).Value

		excelsheets_CF(2).cells(x+1,5) = excel.Cells(mcell.Row, ObjectColumn).Value
		excelsheets_CF(2).cells(x+1,6) = excel.Cells(mcell.Row, ActionColumn).Value
		excelsheets_CF(2).cells(x+1,7) = excel.Cells(mcell.Row, DataColumn).Value
		excelsheets_CF(2).cells(x+1,8) = excel.Cells(mcell.Row, CommentColumn).Value
		excelsheets_CF(2).cells(x+1,9) = mcell.Row

	End If

next

End If



'Browser and Page
If aBrowser <> "" and aPage <>  "" and aObject = "" and Action = "" and Data = "" Then

Set FilterRange =excelsheets(1).UsedRange

FilterRange.Autofilter 1, aBrowser
	FilterRange.Autofilter 2, aPage

RowCount = filterRange.SpecialCells(xlCellTypeVisible).Rows.Count

For each mcell in FilterRange.Specialcells(xlCellTypeVisible).Rows

	If mcell.Row > 1 Then

		  x =GetEmptyRowNumber(excelsheets_CF(2))

		  excelsheets_CF(2).cells(x+1,1) = tst.name
		  excelsheets_CF(2).cells(x+1,2) = tst.id
		  excelsheets_CF(2).cells(x+1,3) = excel.Cells(mcell.Row, BrowserColumn).Value
          excelsheets_CF(2).cells(x+1,4) = excel.Cells(mcell.Row, PageColumn).Value
		  excelsheets_CF(2).cells(x+1,9) = mcell.Row

	End If

next


End If


'Browser, Page, Action, and Data
If aBrowser <> "" and aPage <>  "" and aObject = "" and Action <> "" and Data <> "" Then


Set FilterRange =excelsheets(1).UsedRange


FilterRange.Autofilter 1, aBrowser
FilterRange.Autofilter 2, aPage
FilterRange.Autofilter 4, Action
FilterRange.Autofilter 5, Data

RowCount = filterRange.SpecialCells(xlCellTypeVisible).Rows.Count



For each mcell in FilterRange.Specialcells(xlCellTypeVisible).Rows

	If mcell.Row > 1 Then


		x =GetEmptyRowNumber(excelsheets_CF(2))


		excelsheets_CF(2).cells(x+1,1) = tst.name
		excelsheets_CF(2).cells(x+1,2) = tst.id
		excelsheets_CF(2).cells(x+1,3) = excel.Cells(mcell.Row, BrowserColumn).Value
		excelsheets_CF(2).cells(x+1,4) = excel.Cells(mcell.Row, PageColumn).Value
		excelsheets_CF(2).cells(x+1,6) = excel.Cells(mcell.Row, ActionColumn).Value
		excelsheets_CF(2).cells(x+1,7) = excel.Cells(mcell.Row, DataColumn).Value
		excelsheets_CF(2).cells(x+1,9) = mcell.Row

	End If

next


End If



'Page and Data
If aBrowser = "" and aPage <> "" and aObject = "" and Action = "" and Data <> "" Then

Set FilterRange =excelsheets(1).UsedRange

FilterRange.Autofilter 2, aPage
FilterRange.Autofilter 5, Data

RowCount = filterRange.SpecialCells(xlCellTypeVisible).Rows.Count

For each mcell in FilterRange.Specialcells(xlCellTypeVisible).Rows

	If mcell.Row > 1 Then

		  x =GetEmptyRowNumber(excelsheets_CF(2))

		  excelsheets_CF(2).cells(x+1,1) = tst.name
		  excelsheets_CF(2).cells(x+1,2) = tst.id
		  excelsheets_CF(2).cells(x+1,4) = excel.Cells(mcell.Row, PageColumn).Value
		  excelsheets_CF(2).cells(x+1,7) = excel.Cells(mcell.Row, dataColumn).Value
		  excelsheets_CF(2).cells(x+1,9) = mcell.Row

	End If

next


End If
'===============================================================================
'Object Only
If aBrowser = "" and aPage =  "" and aObject <> "" and Action = "" and Data = "" Then


Set FilterRange =excelsheets(1).UsedRange
FilterRange.Autofilter 3, aObject
RowCount = filterRange.SpecialCells(xlCellTypeVisible).Rows.Count
For each mcell in FilterRange.Specialcells(xlCellTypeVisible).Rows
    If mcell.Row > 1 Then
    x =GetEmptyRowNumber(excelsheets_CF(2))
    excelsheets_CF(2).cells(x+1,1) = tst.name
    excelsheets_CF(2).cells(x+1,2) = tst.id
    excelsheets_CF(2).cells(x+1,3) = excel.Cells(mcell.Row, BrowserColumn).Value
    excelsheets_CF(2).cells(x+1,4) = excel.Cells(mcell.Row, PageColumn).Value
    excelsheets_CF(2).cells(x+1,5) = excel.Cells(mcell.Row, ObjectColumn).Value
    excelsheets_CF(2).cells(x+1,6) = excel.Cells(mcell.Row, ActionColumn).Value
    excelsheets_CF(2).cells(x+1,7) = excel.Cells(mcell.Row, DataColumn).Value
    excelsheets_CF(2).cells(x+1,9) = mcell.Row
    End If
    next
End If

'===============================================================================
'Page and Object only
If aBrowser = "" and aPage <>  "" and aObject <> "" and Action = "" and Data = "" Then
Set FilterRange =excelsheets(1).UsedRange

FilterRange.Autofilter 2, aPage
FilterRange.Autofilter 3, aObject


RowCount = filterRange.SpecialCells(xlCellTypeVisible).Rows.Count
For each mcell in FilterRange.Specialcells(xlCellTypeVisible).Rows
	If mcell.Row > 1 Then
'
	  x =GetEmptyRowNumber(excelsheets_CF(2))
	  excelsheets_CF(2).cells(x+1,1) = tst.name
	  excelsheets_CF(2).cells(x+1,2) = tst.id

	 excelsheets_CF(2).cells(x+1,3) = excel.Cells(mcell.Row, BrowserColumn).Value
	 excelsheets_CF(2).cells(x+1,4) = excel.Cells(mcell.Row, PageColumn).Value
	 excelsheets_CF(2).cells(x+1,5) = excel.Cells(mcell.Row, ObjectColumn).Value
	 excelsheets_CF(2).cells(x+1,6) = excel.Cells(mcell.Row, ActionColumn).Value
	 excelsheets_CF(2).cells(x+1,7) = excel.Cells(mcell.Row, DataColumn).Value
	 excelsheets_CF(2).cells(x+1,8) = excel.Cells(mcell.Row, CommentColumn).Value
	 excelsheets_CF(2).cells(x+1,9) = mcell.Row
	End If

next

End If

'===============================================================================
'Object and Data only
If aBrowser = "" and aPage =  "" and aObject <> "" and Action = "" and Data <> "" Then
Set FilterRange =excelsheets(1).UsedRange

FilterRange.Autofilter 3, aObject
FilterRange.Autofilter 5, Data

RowCount = filterRange.SpecialCells(xlCellTypeVisible).Rows.Count
For each mcell in FilterRange.Specialcells(xlCellTypeVisible).Rows
	If mcell.Row > 1 Then
	 x =GetEmptyRowNumber(excelsheets_CF(2))
	 excelsheets_CF(2).cells(x+1,1) = tst.name
	 excelsheets_CF(2).cells(x+1,2) = tst.id

	 excelsheets_CF(2).cells(x+1,3) = excel.Cells(mcell.Row, BrowserColumn).Value
	 excelsheets_CF(2).cells(x+1,4) = excel.Cells(mcell.Row, PageColumn).Value
	 excelsheets_CF(2).cells(x+1,5) = excel.Cells(mcell.Row, ObjectColumn).Value
	 excelsheets_CF(2).cells(x+1,6) = excel.Cells(mcell.Row, ActionColumn).Value
	 excelsheets_CF(2).cells(x+1,7) = excel.Cells(mcell.Row, DataColumn).Value
	 excelsheets_CF(2).cells(x+1,8) = excel.Cells(mcell.Row, CommentColumn).Value
	 excelsheets_CF(2).cells(x+1,9) = mcell.Row
	End If

next

End If

'======================================================================================================================
'Page, Object and Action only
If aBrowser = "" and aPage <>  "" and aObject <> "" and Action <> "" and Data = "" Then
Set FilterRange =excelsheets(1).UsedRange

FilterRange.Autofilter 2, aPage
FilterRange.Autofilter 3, aObject
FilterRange.Autofilter 4, Action

RowCount = filterRange.SpecialCells(xlCellTypeVisible).Rows.Count
For each mcell in FilterRange.Specialcells(xlCellTypeVisible).Rows
	If mcell.Row > 1 Then

 x =GetEmptyRowNumber(excelsheets_CF(2))
 excelsheets_CF(2).cells(x+1,1) = tst.name
 excelsheets_CF(2).cells(x+1,2) = tst.id

 excelsheets_CF(2).cells(x+1,3) = excel.Cells(mcell.Row, BrowserColumn).Value
 excelsheets_CF(2).cells(x+1,4) = excel.Cells(mcell.Row, PageColumn).Value
 excelsheets_CF(2).cells(x+1,5) = excel.Cells(mcell.Row, ObjectColumn).Value
 excelsheets_CF(2).cells(x+1,6) = excel.Cells(mcell.Row, ActionColumn).Value
 excelsheets_CF(2).cells(x+1,7) = excel.Cells(mcell.Row, DataColumn).Value
 excelsheets_CF(2).cells(x+1,8) = excel.Cells(mcell.Row, CommentColumn).Value
 excelsheets_CF(2).cells(x+1,9) = mcell.Row

	End If

next

End If

'===============================================================================
'Page, Object and Action only
If aBrowser <> "" and aPage <>  "" and aObject <> "" and Action <> "" and Data <> "" Then
Set FilterRange =excelsheets(1).UsedRange

FilterRange.Autofilter 1, aBrowser
FilterRange.Autofilter 2, aPage
FilterRange.Autofilter 3, aObject
FilterRange.Autofilter 4, Action
FilterRange.Autofilter 5, Data

RowCount = filterRange.SpecialCells(xlCellTypeVisible).Rows.Count
For each mcell in FilterRange.Specialcells(xlCellTypeVisible).Rows
	If mcell.Row > 1 Then

	 x =GetEmptyRowNumber(excelsheets_CF(2))
	 excelsheets_CF(2).cells(x+1,1) = tst.name
	 excelsheets_CF(2).cells(x+1,2) = tst.id

	 excelsheets_CF(2).cells(x+1,3) = excel.Cells(mcell.Row, BrowserColumn).Value
	 excelsheets_CF(2).cells(x+1,4) = excel.Cells(mcell.Row, PageColumn).Value
	 excelsheets_CF(2).cells(x+1,5) = excel.Cells(mcell.Row, ObjectColumn).Value
	 excelsheets_CF(2).cells(x+1,6) = excel.Cells(mcell.Row, ActionColumn).Value
	 excelsheets_CF(2).cells(x+1,7) = excel.Cells(mcell.Row, DataColumn).Value
	 excelsheets_CF(2).cells(x+1,8) = excel.Cells(mcell.Row, CommentColumn).Value
	 excelsheets_CF(2).cells(x+1,9) = mcell.Row
	End If

next

End If
'===============================================================================
next


excelwkbk.Close False
excel.Quit
Set excel = nothing
Set excelwkbk= nothing
Set excelsheets = nothing
Set aTestObject = nothing
Set AttachmentFact = nothing
Set AttachmentColl = nothing
Set FilterRange = nothing
End If

If sflag = False Then
'print tst.id
'print tst.name
Print "The " & tst.name & " test doesn't have a valid DataTable.xl" 'Better to write this in a file
'msgbox "The " & tst.name & " file doesn't have a valid DataTable.xl" 'Should be removed as it
	  					'will display too may Message boxes.
Set aTestObject = nothing
Set AttachmentFact = nothing
Set AttachmentColl = nothing
Set theAttach = nothing

End If

    next
Next


Public Function GetTestObjects(tsFolderPath)

Const HasAttachment = "Y"
Const TestType ="QUICKTEST_TEST"
Set tda = qcutil.QCConnection
Set labTreeMgr = tda.TreeManager
Set labFolder = labTreeMgr.NodeByPath(tsFolderPath)
Set testfac = labFolder.TestFactory
Set afilter = testfac.Filter
afilter("TS_ATTACHMENT") = HasAttachment
afilter("TS_TYPE") = TestType
Set tsList = aFilter.NewList()
Set GetTestObjects =tsList
Set tda =nothing
Set labTreeMgr =nothing
Set labFolder = nothing
Set Testfac = nothing


End Function


Function GetDataTableAttachement(tdc,testid)


counter = 0
sflag =False

Set aTestObject = tdc.TestFactory.item(tst.id)
Set AttachmentFact = aTestObject.Attachments
Set AttachmentColl = AttachmentFact.NewList("")
ATT =AttachmentColl.count


For each mfile in AttachmentColl

    counter = counter +1
    mysplit = split(mfile.name,"_",-1)
    myval = lcase(ubound(mysplit))

    If myval = 2 and  (lcase(mysplit (2)))="datatable.xls" Then

        Set theAttach = AttachmentColl.Item(counter)
        theAttach.Load True, ""
        theFileName = theAttach.FileName
        thePath = Left(theFileName, InstrRev(theFileName, "\")-1)
        sflag = True
        Exit for
    End If

Next

If sflag = True Then

    Set excel = createobject("Excel.Application")
    Set excelwkbk = excel.Workbooks.Open(theFileName)
    Set excelwkshts = excelwkbk.Sheets
    excel.Visible =True
    GetDataTableAttachement = True

End If

If sflag = False Then
GetDataTableAttachement = False
End If

Set aTestObject = nothing
Set AttachmentFact = nothing
Set AttachmentColl = nothing
Set theAttach = nothing


End Function


Function GetEmptyRowNumber(sheetobject)

emptyRow =excelsheets_CF(2).UsedRange.Rows.count
GetEmptyRowNumber = emptyRow

End Function

'==========================================================================
'If  If sflag = False Then Then
	'msgbox "The " & tst.name & " file doesn't have a valid DataTable.xl"
'End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
