'#COMMENT
'If used in VBScript / HP UFT uncomment this (see also last part)
'replace <space>as<space> with ' as ... so variable type is gone
'class clsDataTable
'#END COMMENT
'************************************************************************************************************************************************
'* Module       : clsDataTable
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Team         : DEVOPS
'* Version      : 0.8
'************************************************************************************************************************************************
'* Purpose: Having a dataset wrapper thats partly compatible with HP UFT to ease migration
'* Inputs:  N/A
'* Returns: N/A
'************************************************************************************************************************************************
'* Reviewed by  :
'************************************************************************************************************************************************
Private Const mcMaxColumns = 1024
Private m_range As Range
Private m_currentrow As Long
Private m_parameters As Dictionary

'************************************************************************************************************************************************
'* Function     : setData
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Version      : 0.8
'* Purpose      : Sets a reference to the range of data in excel using excel range object
'* Inputs       : n/a
'* Returns      : nothing
'************************************************************************************************************************************************
Sub setData(m_ws As Worksheet, iRowStart, iRowEnd)
    Dim iColumns As Long
    Dim iCol As Long
    
    'Put the field/parameter headers in a dd
    Set m_parameters = createobject("scripting.dictionary")
    
'    iColumn = m_ws.Cells(iRowStart + 1, 1).End(xlToRight).Column + 1
    iColumns = m_ws.Cells(iRowStart + 1, 1).End(-4161).Column + 1 'xlToRight=-4161
    
    For iCol = 1 To iColumns
        m_parameters.Add LCase(m_ws.Cells(iRowStart + 1, iCol)), iCol
    Next
    
    'When its an UIA 1.0 datasheet
    If iRowStart = 0 Then
        Set m_range = m_ws.Range(m_ws.Cells(1, 1), m_ws.Cells(iRowEnd, iColumns))
    Else
        Set m_range = m_ws.Range(m_ws.Cells(iRowStart, 1), m_ws.Cells(iRowEnd, iColumns))
    End If
    
End Sub
'************************************************************************************************************************************************
'* Function     : value
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Version      : 0.8
'* Purpose      : returns a value based on the current row and first row having fieldname, row 1 is the datatable marker line
'* Inputs       : n/a
'* Returns      : nothing
'************************************************************************************************************************************************
Function Value(strParameterName As String) As Variant
    Dim iColumn As Long
    
    If m_parameters.exists(LCase(strParameterName)) Then
        iColumn = m_parameters.Item(LCase(strParameterName))
    End If

    Value = m_range.Cells(m_currentrow, iColumn)
End Function
'************************************************************************************************************************************************
'* Function     : getRowCount
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Version      : 0.8
'* Purpose      : returns the number of rows
'* Inputs       : n/a
'* Returns      : nothing
'************************************************************************************************************************************************
Function getRowcount() As Long
    getRowcount = m_range.Rows.Count
End Function
Function getParameterCount() As Long
    getParameterCount = m_range.Columns.Count
End Function
Function getRange() As Range
    Set getRange = m_range
End Function
Sub setCurrentRow(iRow As Long)
    m_currentrow = iRow + 1
End Sub
'************************************************************************************************************************************************
'* Function     : getAsArray
'* Date         : May 2018
'* Author       : Elwin Wildschut
'* Version      : 0.8
'* Purpose      : Returns an array
'* Inputs       : n/a
'* Returns      : n/a
'************************************************************************************************************************************************
Function getAsArray()
    Dim Arr As Variant
    Arr = m_range
    getAsArray = Arr
End Function

'#COMMENT
'If used in VBScript / HP UFT uncomment this (see also last part)
'replace <space>as<space> with ' as ... so variable type is gone
'end class
'Function new_clsDataTable()
'    Set new_clsDataTable = New clsDataTable
'End Function
'#END COMMENT
