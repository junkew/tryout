Option Explicit
'#COMMENT
'If used in VBScript / HP UFT uncomment this (see also last part)
'replace <space>as<space> with ' as ... so variable type is gone
'class clsDataTables
'#END COMMENT
'************************************************************************************************************************************************
'* Module       : clsDataTables
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Team         : DEVOPS
'* Version      : 0.6
'************************************************************************************************************************************************
'* Purpose: Having a datamanager to manage the datasets (which could be in xls, mdb, csv, ....)
'*                  Only managing now XLS files and scans all worksheets
'* Inputs:  N/A
'* Returns: N/A
'************************************************************************************************************************************************
'* Reviewed by  :
'************************************************************************************************************************************************
Private m_ac As Object                          'Reference to application Activator context if its not used indepently
Private m_xlsApp As excel.Application           'Reference to Excel application
Private m_Wb As Workbook                        'Reference to workbook where we look in for datatables
Private m_ws As Worksheet                       'Reference to worksheet where we look in for datatables normally sheet 1
Private m_dataLocation As String                'Foldername, if set then all workbooks with extensions .xls will be scanned for datatables
Private Const mcDatabookColumn = "A:B"          'Columns we look in for searching a certain dataset (UIA 1.0 has it in column A and new one in column B)
'************************************************************************************************************************************************
'* Function     : main (to stay in Java / .NET logic of main routine as alternative to class_initialize
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Version      : 0.8
'* Purpose      : Sets a reference where all objects can be attached to so plugin(s) model can be used
'* Inputs       : n/a
'* Returns      : initialized object
'************************************************************************************************************************************************
Sub main()
    'If this function not exist then the class will act independently together with clsDataTable
    'otherwise it expects an overall framework to have a reference function setApplicationActivator
    If getRef("setApplicationActivator") = True Then
        Set m_ac = setApplicationActivator("Automation manager")
        Set m_xlsApp = m_ac.getExcel
    Else
        Set m_xlsApp = createobject("Excel.application")
        m_xlsApp.Visible = True
    End If
End Sub
'Just dispatch to main
Private Sub Class_Initialize()
    Me.main
End Sub
'************************************************************************************************************************************************
'* Function     : datalocation
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Version      : 0.8
'* Purpose      : Return the datalocation where all dUIAiles should be scanned
'* Inputs       : n/a
'* Returns      : folder
'************************************************************************************************************************************************
Public Property Get dataLocation()
    dataLocation = m_dataLocation
End Property
'************************************************************************************************************************************************
'* Function     : datalocation
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Version      : 0.8
'* Purpose      : Set the datalocation where all dUIAiles should be scanned
'* Inputs       : String to folder location
'* Returns      : nothing
'************************************************************************************************************************************************

Public Property Let dataLocation(ByVal strFolderLocation)
    m_dataLocation = strFolderLocation
End Property
'************************************************************************************************************************************************
'* Function     : getDataset
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Version      : 0.8
'* Purpose      : Returns a reference to a dataset
'* Inputs       : name of the dataset we are looking for
'* Returns      : object of clsDataTable
'************************************************************************************************************************************************
Function getDataSet(strTableName As String) As clsDataTable
    Dim bTableFound As Boolean    'Did we find a table then do not search further
    Dim iRow As Long
    Dim foundCell As Range
    Dim iRowStart As Long
    Dim iColumn As Long
    Dim oDatatable As clsDataTable
    Dim tWs As Worksheet
    
    'If its a single workbook (where most teams start with only 1) we behave differently
    'So in this case end user has called openDatabook
    If Me.dataLocation = "" Then
            
            'Look first for the data in the opened book and build the range rectangular area
            Set foundCell = m_ws.Columns(mcDatabookColumn).Find(strTableName)
            
            If Not foundCell Is Nothing Then
                'Column mcDatabookColumn cannot be empty till row where the range ends
                iRowStart = foundCell.Row
                iRow = foundCell.End(-4121).Row 'xldown=-4121 done this way to be vbscript compatible
                
                Set oDatatable = New clsDataTable
                oDatatable.setData m_ws, iRowStart, iRow
                Set getDataSet = oDatatable
                bTableFound = True
            End If
            
            'Try UIA1.0 databook
            If Not bTableFound = True Then
                'If its an UIA databook then names of the worksheets are the names of the tables
                On Error Resume Next
                Set tWs = m_Wb.Worksheets(strTableName)
                Err.Clear
                On Error GoTo 0
                
                'So when worksheet exist
                If IsObject(tWs) And Not tWs Is Nothing Then
                    iRowStart = 0 'Start on 0, clsDataTable was initially based on a databook where row 1 is reserved
                    iRow = tWs.Cells(1, 1).End(-4121).Row 'xldown=-4121 done this way to be vbscript compatible
                    
                    Set oDatatable = New clsDataTable
                    oDatatable.setData tWs, iRowStart, iRow
                    Set getDataSet = oDatatable
                    bTableFound = True
                Else
                    Err.Clear
                End If
                On Error GoTo 0
            End If
            
    Else
        
    End If
    
End Function
'************************************************************************************************************************************************
'* Function     : openDatabook
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Version      : 0.8
'* Purpose      : Returns a reference to a databook that can hold multiple datasets
'* Inputs       : full filename including foldername
'* Returns      : open workbook reference
'************************************************************************************************************************************************
Sub openDatabook(strFileName As String)
    Set m_Wb = m_xlsApp.workbooks.Open(strFileName)
    Set m_ws = m_Wb.Worksheets(1)
End Sub
'************************************************************************************************************************************************
'* Function     : closeDatabook
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Version      : 0.8
'* Purpose      : Closes a previously opened workbook
'* Inputs       : n/a
'* Returns      : open workbook reference
'************************************************************************************************************************************************
Sub closeDatabook()
    If IsObject(m_Wb) = True Then
        Set m_ws = Nothing
        m_Wb.Close
        Set m_Wb = Nothing
    End If
End Sub
Private Sub Class_Terminate()
    On Error Resume Next 'Most likely we do not want to catch errors anymore when this happens
    closeDatabook
End Sub

'#COMMENT
'If used in VBScript / HP UFT uncomment this (see also last part)
'replace <space>as<space> with ' as ... so variable type is gone
'end class
'Function new_clsDataTables()
'    Set new_clsDataTables = New clsDataTables
'End Function
'#END COMMENT

