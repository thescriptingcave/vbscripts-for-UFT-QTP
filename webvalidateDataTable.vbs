' **************************************************************************************************
'	This Utility/Action  Performs Object Validation. It  verifies the objects ( Browser, Page and Object)  in the data table matches the objects in the Repository .
'   It assumes Valid data in the datatable
'  It searches obejcts from their corrosponing parent calss.
'
' **************************************************************************************************
'Option Explicit

Dim myRep
Dim  browserCollection, pageCollection, objectCollection
Dim rowCount, browserCollectionCount, pageCollectionCount, objectCollectionCount
Dim excelBrowserName, excelPageName, excelObjectName, excelActionName
Dim mySheet, objFound, browserTestObject, pageTestObject, objectTestObject
Dim browserLogicalName, pageLogicalName, objectLogicalName,i,j,k,m
Dim browserfound,pagefound, objectFound


datatableName= environment.Value("DataTableName")
datatable.Import datatableName

'DataTable.GetSheet(1).AddParameter "DataValidation",""    'Adds a column to point to rows that needs data verification
DataTable.GetSheet(1).AddParameter "ORValidation", ""	'Adds a column to point to rows that  consists of  objects not found in the repository

'Removes all shared Object Repositories associated with this Action
RepositoriesCollection.RemoveAll
Set myRep = CreateObject("Mercury.ObjectRepositoryUtil")



'Copies Repository from QC to Local Drive
'OrSource =QCGetResource("OR_CCCW_ALLPages.tsr", "C:\Temp","OR_CCCW_ALLPages.tsr")	'function call
OrSource =QCGetResource("flight2.tsr", "C:\Reports","flight2.tsr")		'Function call. The return value is the full  path of the repository copied to the local drive from the QC
'OrSource ="c:\Reports\flight2.tsr"		'THIS NEEDS TO BE WORKED OUT

myRep.Load OrSource

Set browserCollection = myRep.GetAllObjectsByClass("Browser")
browserCollectionCount = browserCollection.Count

Set mySheet = Datatable.GetSheet(1)

rowCount = mySheet.GetRowCount

'loop through all the rows and validate if the .....
For i = 1 to rowCount
mySheet.SetCurrentRow i
excelActionName = ucase(mySheet.GetParameter("Action"))

'<<<<<<<<<<Not Applicable to the current case<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<,
'If  excelActionName = "RESTART" Then 'If action name is Restart go to the next row since data/value is avaialble in the Restart row
	'i = i+1
	'mySheet.SetCurrentRow i
'End If
'
'Read values from the current row of the datatable
excelActionName = ucase(mySheet.GetParameter("Action"))				'Get the value in the Action  column of the Data Table and assign it  excelActionName
					'print excelActionName
excelBrowserName = mySheet.GetParameter("Browser")						'Get the value in the Browser  column of the Data Table and assign it to the variable called excelBrowserName
					'print excelBrowserName
excelPageName = mySheet.GetParameter("Page")								'Get the value in the Page  column of the Data Table and assign it  to the the variable called excelPageName
					'print excelPageName
excelObjectName = mySheet.GetParameter("Object")							'Get the value in the Object  column of the Data Table and assign it  to the browser calleds excelObjectName
											'print excelObjectName

browserfound = False
pagefound = False
objectFound = False

'################################################################################
'If the action name is diffrent from Restart, Forward, and Backward it verifies that the objects (Browser, Page and Object)  in the data table match with the objects in the Repository


For j = 0 To browserCollectionCount - 1
Set browserTestObject = browserCollection.Item(j)
browserLogicalName = myRep.GetLogicalName(browserTestObject)
If  browserLogicalName = excelBrowserName Then
    browserfound = True
    Set  pageCollection = myRep.GetChildren(browserTestObject)
    pageCollectionCount = pageCollection.count
        For k = 0 To pageCollectionCount  - 1
        Set pageTestObject	=  pageCollection.Item(k)
        pageLogicalName = myRep.GetLogicalName(pageTestObject)
            If pageLogicalName = excelPageName Then
                pagefound = true
                Set objectCollection = myRep.GetChildren(pageTestObject)
                objectCollectionCount = objectCollection.count
                For m = 0 To objectCollectionCount - 1
                Set objectTestObject = objectCollection.Item(m)
                objectLogicalName = myRep.GetLogicalName(objectTestObject)
                If   ucase(objectLogicalName) = ucase(excelobjectName) Then
                objectfound = true
                            Exit for 			'.......... m
                End if
        Next '..........m
If (objectfound = true) then
 Reporter.ReportEvent micPass, "Object", "Object "&excelObjectName&" is found in the repository"
Else
Datatable("ORValidation") = "Object :-" &excelObjectName& "  is  missing in the Repository "
Reporter.ReportEvent micFail, "Object", "Row " & i + 1 & "       Object      "&excelObjectName&"     missing in the Repository for action  Keyword    "&excelActionName

End If
Exit for '......k
End If ' ........k
Next   '............ k
If   		 (pagefound = true) then
'Reporter.ReportEvent micPass, "Object", "Page      "&excelpageName&"     is found in the repository"
Reporter.ReportEvent micPass, "Page Object", "Page      "&excelpageName&"     is found in the repository"
Else
Datatable("ORValidation") = "Page object :-" &excelPageName& "  is  missing in the Repository "
'Reporter.ReportEvent micFail, "Object", "object      "&excelpageName&"     missing in the Repository for action  Keyword    "
Reporter.ReportEvent micFail, "Page Object", "Row " & i + 1& "      Page      "&excelpageName&"     missing in the Repository for action  Keyword    "
								End If

								Exit for '.....j
					End if '.......j
Next  ''...........j

If    (browserfound = true) then
'Reporter.ReportEvent micPass, "Object", "Browser      "&excelbrowserName&"     is found in the repository"
Reporter.ReportEvent micPass, "Browser Object", "Browser      "&excelbrowserName&"     is found in the repository"
Else
'Datatable("ORValidation") = "Browser :-" &excelBrowserName&"  is  missing in the Repository "
Datatable("ORValidation") = "Browser Object:-" &excelBrowserName&"  is  missing in the Repository "
'Reporter.ReportEvent micFail, "Object", "Browser      "&excelbrowserName&"     missing in the Repository for action  Keyword    "
Reporter.ReportEvent micFail, "Browser Object", "Row " & i + 1& "        Browser      "&excelbrowserName&"     missing in the Repository for action  Keyword    "
End if

Next



'###############################################################################'#########################

Function QCGetResource(in_resource, dest_folder, dest_file)

	Dim oQTP, oQC, oRes, oFilter, oFile, sDestination, sName, oFileList

	Set oQTP = CreateObject("QuickTest.Application")
	Set oQC = QCUtil.QCConnection
	Set oRes = oQC.QCResourceFactory
	Set oFilter =  oRes.Filter

	sDestination = dest_folder
	sName = dest_file

	oFilter.Filter("RSC_FILE_NAME") = sName

	Set oFileList = oFilter.NewList

	If oFileList.Count = 1 Then
		Set oFile = oFileList.Item(1)
		oFile.FileName = in_resource
		oFile.DownloadResource sDestination, True
	End If

	Set oQTP = Nothing
	Set oQC = Nothing
	Set oRes = Nothing
	Set oFilter = Nothing
	Set oFileList = Nothing
	Set oFile = Nothing

	QCGetResource = dest_folder & "\" & dest_file

End Function

'export the run time datatable
'Datatable.Export  "C:\DataTables\Testdata_afterexec.xls"
Datatable.Export  "C:\DataTables\Validated Web Datatable.xls"

'MsgBox "Test Complete!! Look @ C:\DataTables\Testdata_afterexec.xls File to check Reported Errors"
'MsgBox "Test Completed!! Look @ C:\DataTables\Validated Web Datatable.xls File to check Reported Errors"


'Set browserCollection = Nothing
'Set mySheet = Nothing
'Set myRep = Nothing
