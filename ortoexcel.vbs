
Option Explicit

Dim ExcelApp, Sheet1, ORCopy, MyFile, qtApp
Dim objWorkBook, RepositoryFrom, ORSource
Dim browserCollection, browserLogicalName, browserTestObject
Dim objPageObject, pageCollection, pageTestObject, pageLogicalName
Dim objectCollection, objectTestObject, objectLogicalName
Dim rc, results_filename, results_path, datatableName, window
Dim intLen, getstr,  itemCnt , char
Dim i, k, m, j, fso
Dim ScreenName

j = 0

results_path = "c:\temp"
Set fso = CreateObject("Scripting.FileSystemObject")
Set objPageObject = Nothing

Set RepositoryFrom = Nothing
Set RepositoryFrom = CreateObject("Mercury.ObjectRepositoryUtil")

'Remove all Shared Object Repositories associated with this action
RepositoriesCollection.RemoveAll()
'OrSource = QCGetResource("flight2.tsr", "C:\Reports","flight2.tsr")

RepositoryFrom.Load "C:\Temp\web.tsr"


window = "DisabilityInsurance_ContactSDI"

'window = environment.Value("ScreenName")                  '               Passed from VB script

intLen = Len(window)
getstr = ""

For itemCnt = 1 To intLen
char = Mid(window, itemCnt, 1)
If (char = ":") Then
        getstr = getstr & "_"
		print getstr
Else
        getstr = getstr + char
		print getstr
End If
Next

Dim objWorksheet
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible = False
'ExcelApp.Visible = True                                                                                ' Used for debugging
Set objWorkBook = ExcelApp.Workbooks.Add()
Set objWorksheet = objWorkBook.Worksheets(1)
objWorksheet.Name = getstr                                                    'Renames the worksheet

objWorksheet.cells (1, "A").ColumnWidth = 60


Set browserCollection = RepositoryFrom.GetAllObjectsByClass("Browser")
                                                                                'print browserCollection.Count
For i = 0 to browserCollection.count - 1
Set browserTestObject = browserCollection.item(i)
browserLogicalName = RepositoryFrom.GetLogicalName(browserTestObject)
                                                                                'print browserLogicalName                                                                          'Because there is only one browser object, the for loop will be excuted only once
                                                                                'assumes that there is only one browser, else see validate data table on how to handle more than one browser
Set pageCollection = RepositoryFrom.GetChildren(browserTestObject)
                                                                                'print "Number of Pages: " & pageCollection.count
For k = 0 to pageCollection.count - 1
Set pageTestObject = pageCollection.item(k)
pageLogicalName = RepositoryFrom.GetLogicalName(pageTestObject)
                                                                                'print "Page Logical name: " & pageLogicalName
If  pageLogicalName = window Then
                j =  j + 1                               'The value of j does not matter. Page is in the OR
                                                                                Set objectCollection = RepositoryFrom.GetAllObjects(pageTestObject)
                                                                                                                                                                'print "Number of objects under " & pageLogicalName & " is " & objectCollection.count
                                                                                For m = 0 To objectCollection.count - 1
                                                                                Set  objectTestObject= objectCollection.Item(m)
                                                                                objectLogicalName=repositoryFrom.GetLogicalName(objectTestObject)                                                                                               'print objectLogicalName
                                                                                objWorksheet.cells(m +1, "A") =  objectLogicalName
                                                                                            Next
                                                                                        End If
                                                                                Next
Next

If j = 0 Then
'Page is not in the OR
Msgbox "Page " & window & " is NOT in the OR"
ExitTest
End If

'Termination rc
Termination
'ExitAction (rc)


'Private Function Termination (rc)
Private Function Termination
On Error Resume Next
results_filename = results_path & "\" & getstr & ".xls"
fso.DeleteFile results_filename
wait (1)
On Error GoTo 0
If (fso.FolderExists(results_path)) Then
    Else
    Set nf = fso.CreateFolder(results_path)
End If
'rc= objWorkBook.SaveAs (results_filename)
objWorkBook.SaveAs (results_filename)
objWorkBook.Close True
ExcelApp.Quit
Set ExcelApp = Nothing
Set fso = Nothing
MsgBox "All Done!" & chr(10) & chr(13) & "Results are in file: " &              results_filename
End Function
