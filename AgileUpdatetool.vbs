'Update Tool developed by Robert Harding, Robert Harding Consulting

Const str = "c:\temp\controlfile.xlsx"
Const xlCellTypeVisible =12
Const HasAttachment = "Y"
Const TestType ="QUICKTEST_TEST"


const browser_column =1
const page_column = 2
const object_column = 3
const action_column =4
const data_column = 5
'Const step_column =6

const Excel2003fileformat = 56

Set x = createobject("excel.application")
Set y = x.Workbooks.Open(str)
Set z = y.Worksheets
Set listofTest = createobject("Scripting.Dictionary")
x.Visible = true
Set cf = z(2)
spos_counter =3
Attachmentflag =True

Set RangeAllRowsColumns = cf.UsedRange
row_totals_of_range_RangeAllRowsColumns = RangeAllRowsColumns.Rows.Count
For i = 2 to row_totals_of_range_RangeAllRowsColumns
spos  = RangeAllRowsColumns(spos_counter,1)
fpos = RangeAllRowsColumns(i,1)

    if fpos <> spos Then
    listofTest.Add i, RangeAllRowsColumns(i,1)
    firstpos = secondpos
    End If

spos_counter = spos_counter + 1

Next

For each test_name in ListofTest.items

print test_name

cf.UsedRange.AutoFilter 1, test_name
Set objTarget = cf.UsedRange.find(test_name,,,1)
 rowpos = objTarget.Row
 colpos = objTarget.Column
 test_object_id  = cint(cf.Cells(rowpos, colpos + 1))
 print test_object_id

'***********************************************************************************

Set test_object= qcutil.QCConnection.TestFactory.item(test_object_id) 'pass the test object ID
TestCheckedStatus = QC_VCSCheckOutTest(test_object)

If TestCheckedStatus <> "" Then
'checkoutflag = True
' msgbox "Please CHECK-IN test " & test_object.id & " FIRST and then Select the OK button"
msgbox "Please CHECK-IN test " & test_object.id & " and other CHECKED-OUT tests FIRST. " & chr(10) & chr(13) & _
chr(10) & chr(13) &  "Data Table Updated for all CHECKED-IN tests before test " & test_object.id & " , if any, in the control file"
ExitTest

End If
Set attachmentfactory = test_object.Attachments
Set aFilter = attachmentfactory.Filter
aFilter("TS_ATTACHMENT") = HasAttachment
aFilter("TS_TYPE") = TestType
Set attachmentList = attachmentfactory.NewList("")
If attachmentList.Count = 0 Then
       Attachmentflag =false
End If

For each theattachment in attachmentList
tn =ucase("Test_") & test_object.id & "_" & ucase("datatable.xls")
 If ucase(theattachment.name) =ucase(tn) Then
        Exit FOR
 END IF
Next

filepath = theattachment.FileName
fileName = theattachment.name
filepathlength =len(filepath)
pos =instr(1,filepath,fileName)
dirpath =trim(mid(filepath,1,pos -1))
newdestination = dirpath & "datatable.xls"

    '**************************************************************************************
Set a = createobject("excel.application")
Set b = a.Workbooks.Open(filepath)
Set c = b.Worksheets
a.Visible = True

If attachmentflag = false Then
'msgbox "There are no attachments for this test id, verify that the attachment exist" 'stop
msgbox "Check if test with test id " & test_object_id & " exists and has also an attachment"
ExitTest
End If

row_count = cf.UsedRange.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1

print "total number of rows for this Test = " & row_count
			' print "Row Count is " & row_count


set rngfiltered = cf.UsedRange.Columns(1).SpecialCells(xlCellTypeVisible).Rows

For each filtervalue in rngfiltered.Rows
If filtervalue.row > 1 Then
    row_number = cf.Cells(filtervalue.row, 9)
    BrowserUpdate =cf.Cells(filtervalue.row, 10)
    PageUpdate =cf.Cells(filtervalue.row, 11)
    ObjectUpdate =cf.Cells(filtervalue.row, 12)
    ActionUpdate =cf.Cells(filtervalue.row, 13)
    DataUpdate = cf.Cells(filtervalue.row, 14)

If  BrowserUpdate <> "" or PageUpdate  <> "" or ObjectUpdate <> "" or           ActionUpdate  <> "" or  DataUpdate <> "" then
      'print " row number  " & filtervalue.row  & " will be updated" & vblf
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Browser Only
If BrowserUpdate <> "" and PageUpdate =""  and ObjectUpdate = "" and ActionUpdate ="" and DataUpdate = "" Then
    c(1).Cells(row_number,browser_column) =BrowserUpdate
    c(1).Cells(row_number,browser_column).interior.ColorIndex =36
End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Object Only

If BrowserUpdate = "" and PageUpdate =""  and ObjectUpdate <> "" and ActionUpdate ="" and DataUpdate = "" Then
  c(1).Cells(row_number,object_column) = ObjectUpdate
  c(1).Cells(row_number,object_column).interior.ColorIndex =36
End If

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Page Only
If BrowseUpdate = "" and PageUpdate <>"" and ObjectUpdate = "" and ActionUpdate = "" and DataUpdate = "" Then
    c(1).Cells(row_number,page_column) = PageUpdate
	c(1).Cells(row_number,page_column).interior.ColorIndex =36
End If

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Data Only
If BrowserUpdate = "" and PageUpdate ="" and ObjectUpdate = "" and ActionUpdate = "" and DataUpdate <> "" Then
      mstring = "Data value changed from " & data & " to " & DataUpdate
      c(1).Cells(row_number,data_column) =DataUpdate
      c(1).Cells(row_number,data_column).interior.ColorIndex =36
End If

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Browser and Page
If BrowserUpdate <> "" and PageUpdate <>"" and ObjectUpdate = "" and ActionUpdate = "" and  DataUpdate = "" Then
		  c(1).Cells(row_number,browser_column) =BrowserUpdate
          c(1).Cells(row_number,browser_column).interior.ColorIndex =36
          c(1).Cells(row_number,page_column) = PageUpdate
          c(1).Cells(row_number,page_column).interior.ColorIndex =36

End If

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Browser and Object
If BrowserUpdate <> "" and PageUpdate = "" and ObjectUpdate <> "" and ActionUpdate = "" and  DataUpdate = "" Then
          c(1).Cells(row_number,browser_column) =BrowserUpdate
       	  c(1).Cells(row_number,browser_column).interior.ColorIndex =36
       	  c(1).Cells(row_number,object_column) = ObjectUpdate
       	  c(1).Cells(row_number,object_column).interior.ColorIndex =36

End If

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Page and Object
If BrowserUpdate = "" and PageUpdate <> "" and ObjectUpdate <> "" and ActionUpdate = "" and  DataUpdate = "" Then
          c(1).Cells(row_number,page_column) = PageUpdate
          c(1).Cells(row_number,page_column).interior.ColorIndex =36
          c(1).Cells(row_number,object_column) = ObjectUpdate
          c(1).Cells(row_number,object_column).interior.ColorIndex =36
End If

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Object and Data
If BrowserUpdate = "" and PageUpdate = "" and ObjectUpdate <> "" and ActionUpdate = "" and  DataUpdate <> "" Then
          c(1).Cells(row_number,object_column) = ObjectUpdate
          c(1).Cells(row_number,object_column).interior.ColorIndex =36
		  c(1).Cells(row_number,data_column) = DataUpdate
          c(1).Cells(row_number,data_column).interior.ColorIndex =36

End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Browser and Data
If BrowserUpdate <> "" and PageUpdate ="" and ObjectUpdate = ""  and ActionUpdate = "" and  DataUpdate <> "" Then

		c(1).Cells(row_number,browser_column)= BrowserUpdate
        c(1).Cells(row_number,browser_column).interior.ColorIndex =36
        c(1).Cells(row_number,data_column)=DataUpdate
        c(1).Cells(row_number,data_column).interior.ColorIndex =36
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Page and Data
If BrowserUpdate = "" and PageUpdate <>"" and ObjectUpdate = "" and ActionUpdate = "" and  DataUpdate <> "" Then
        c(1).Cells(row_number,page_column)= PageUpdate
        c(1).Cells(row_number,page_column).interior.ColorIndex =36
        c(1).Cells(row_number,data_column)=DataUpdate
        c(1).Cells(row_number,data_column).interior.ColorIndex =36
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Browser, Page, Data
If BrowserUpdate <> "" and PageUpdate <>"" and ObjectUpdate = "" and ActionUpdate = "" and  DataUpdate <> "" Then

		 c(1).Cells(row_number,browser_column) =BrowserUpdate
     	 c(1).Cells(row_number,browser_column).interior.ColorIndex = 36
         c(1).Cells(row_number,page_column) = PageUpdate
         c(1).Cells(row_number,page_column).interior.ColorIndex = 36
         c(1).Cells(row_number,data_column) =DataUpdate
         c(1).Cells(row_number,data_column).interior.ColorIndex = 36

End If

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Browser, Page, Object
If BrowserUpdate <> "" and PageUpdate <>"" and ObjectUpdate <> "" and ActionUpdate = "" and  DataUpdate = "" Then
		  c(1).Cells(row_number,browser_column) =BrowserUpdate
          c(1).Cells(row_number,browser_column).interior.ColorIndex = 36
          c(1).Cells(row_number,page_column) = PageUpdate
          c(1).Cells(row_number,page_column).interior.ColorIndex = 36
          c(1).Cells(row_number,object_column) = ObjectUpdate
          c(1).Cells(row_number,object_column).interior.ColorIndex = 36

End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Browser, Object and Data
If BrowserUpdate <> "" and PageUpdate = "" and ObjectUpdate <> "" and ActionUpdate = "" and  DataUpdate <> "" Then
         c(1).Cells(row_number,browser_column) =BrowserUpdate
         c(1).Cells(row_number,browser_column).interior.ColorIndex = 36
         c(1).Cells(row_number,object_column) = ObjectUpdate
         c(1).Cells(row_number,object_column).interior.ColorIndex = 36
         c(1).Cells(row_number,data_column) = DataUpdate
         c(1).Cells(row_number,data_column).interior.ColorIndex = 36

End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Page, Object and Data
If BrowserUpdate = "" and PageUpdate <>"" and ObjectUpdate <> "" and ActionUpdate = "" and  DataUpdate <> "" Then
		 c(1).Cells(row_number,page_column) = PageUpdate
         c(1).Cells(row_number,page_column).interior.ColorIndex = 36
	     c(1).Cells(row_number,object_column) = ObjectUpdate
         c(1).Cells(row_number,object_column).interior.ColorIndex = 36
         c(1).Cells(row_number,data_column) =DataUpdate
         c(1).Cells(row_number,data_column).interior.ColorIndex = 36

End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 'Object, Action and Data
 If BrowserUpdate = "" and PageUpdate = "" and ObjectUpdate <> "" and ActionUpdate <> "" and  DataUpdate <> "" Then
		 c(1).Cells(row_number,object_column) = ObjectUpdate
         c(1).Cells(row_number,object_column).interior.ColorIndex = 36
	     c(1).Cells(row_number,action_column) = ActionUpdate
         c(1).Cells(row_number,action_column).interior.ColorIndex = 36
         c(1).Cells(row_number,data_column) =DataUpdate
         c(1).Cells(row_number,data_column).interior.ColorIndex = 36

End If

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Browser, Page, Object, and Data
If BrowserUpdate <> "" and PageUpdate <>"" and ObjectUpdate <> "" and  ActionUpdate = "" and  DataUpdate <> "" Then
		c(1).Cells(row_number,browser_column) = BrowserUpdate
        c(1).Cells(row_number,browser_column).interior.ColorIndex =36
        c(1).Cells(row_number,page_column) = PageUpdate
        c(1).Cells(row_number,page_column).interior.ColorIndex =36
		c(1).Cells(row_number,object_column) = ObjectUpdate
        c(1).Cells(row_number,object_column).interior.ColorIndex =36
        c(1).Cells(row_number,data_column) =DataUpdate
	    c(1).Cells(row_number,data_column).interior.ColorIndex =36

End If
 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Browser, Page, Object, Action and Data
If BrowserUpdate <> "" and PageUpdate <>"" and ObjectUpdate <> "" and  ActionUpdate <> "" and  DataUpdate <> "" Then
		c(1).Cells(row_number,browser_column) = BrowserUpdate
        c(1).Cells(row_number,browser_column).interior.ColorIndex =36
        c(1).Cells(row_number,page_column) = PageUpdate
        c(1).Cells(row_number,page_column).interior.ColorIndex =36
		c(1).Cells(row_number,object_column) = ObjectUpdate
        c(1).Cells(row_number,object_column).interior.ColorIndex =36
		c(1).Cells(row_number,action_column) = ActionUpdate
	    c(1).Cells(row_number,action_column).interior.ColorIndex =36
        c(1).Cells(row_number,data_column) =DataUpdate
	    c(1).Cells(row_number,data_column).interior.ColorIndex =36

End If

 '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        End If
    End If
Next

b.Save
a.Quit


theattachment.Save True
theattachment.post

Set excelapp = nothing
Set excelwkbk = nothing
Set fso = createobject("scripting.FileSystemObject")
fso.CopyFile filepath, newdestination ,True

Set Vercontrolobject = nothing
Set Vercontrolobject = test_object.vcs

Vercontrolobject.Refresh
Vercontrolobject.CheckOut -1, "",False

theattachment.Rename("datatable_automation" & second(Now) &".xls")
SaveAttachmentToTestObj test_object,newdestination ,"Updated by Automation"

Vercontrolobject.Checkin "","Updated By Automation"

Next

msgbox "Data Table Updated"			'Just included see if it helps

Public Sub SaveAttachmentToTestObj(TestObj, LocalFilePath, FileDescription)
    Set Attachments = TestObj.Attachments
    Set Attachment = Attachments.AddItem(Null)
   Attachment.FileName = LocalFilePath
                                        'c(1).Cells(row_number,action_column) =ActionUpdate 	'We don't need Action
    Attachment.Description = FileDescription
    Attachment.Type = 1'TDATT_FILE
    Attachment.Post ' Commit changes
End Sub

Public Function QC_VCSCheckOutTest(testobject)
' IsTestCheckOut returns blank if it is not checked out, return the person's name that has it checked out
Dim VersionControlObject, VCS_CheckOut
Set VersionControlObject = testobject.Vcs
VersionControlObject.Refresh
IsTestCheckedOut = VersionControlObject.Lockedby
QC_VCSCheckOutTest = IsTestCheckedOut
End Function
