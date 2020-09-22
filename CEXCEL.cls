VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEXCEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------
' CEXCEL Class developed by Clint LaFever
'
' http://vbasic.iscool.net
'------------------------------------------------------------
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public XL As Object
Private ExcelWasNotRunning As Boolean
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Open the passed workbook with Excel, Hidden or
'                Visible
' Date: November,27 2001 @ 13:32:47
'------------------------------------------------------------
Public Sub OpenExcel(Filename As String, Optional MakeVisible As Boolean = False)
    On Error Resume Next
    Set XL = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear
    DetectExcel
    Set XL = GetObject(Filename)
    XL.Application.Visible = MakeVisible
    XL.Parent.Windows(1).Visible = True
    End Sub
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Close excel and either save or dont
' Date: November,27 2001 @ 13:32:57
'------------------------------------------------------------
Public Sub CloseExcel(SaveWorkBooks As Boolean)
    Dim w, x
    For Each x In XL.Parent.Windows
        x.Visible = True
    Next
    If SaveWorkBooks = True Then
        For Each w In XL.Application.Workbooks
            w.Save
        Next w
    Else
        XL.Application.DisplayAlerts = False
    End If
    If ExcelWasNotRunning = True Then
        XL.Application.Quit
    End If
    Set XL = Nothing
End Sub
Private Sub DetectExcel()
    Const WM_USER = 1024
    Dim hwnd As Long
    hwnd = FindWindow("XLMAIN", 0)
    If hwnd = 0 Then
        Exit Sub
    Else
        SendMessage hwnd, WM_USER + 18, 0, 0
    End If
End Sub


'------------------------------------------------------------
' The below function pretty much are just shortcuts
' to VBA code.  As you see, all of them use the
' XL object which you can just use too if you want.
'------------------------------------------------------------




'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Get the number of sheets
' Date: November,27 2001 @ 13:33:12
'------------------------------------------------------------
Public Function SheetCount() As Long
    On Error Resume Next
    SheetCount = XL.Sheets.Count
End Function
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Set the name of a sheet
' Date: November,27 2001 @ 13:33:21
'------------------------------------------------------------
Public Function SheetName(sNUM As Long) As String
    On Error Resume Next
    SheetName = XL.Sheets(sNUM).Name
End Function
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Get the contents of a cell
' Date: November,27 2001 @ 13:33:31
'------------------------------------------------------------
Public Function GetCell(SheetNumber As Long, Column As String, row As Long) As Variant
    On Error Resume Next
    GetCell = XL.Sheets(SheetNumber).Range(Column & row)
End Function
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Write to a cell
' Date: November,27 2001 @ 13:33:36
'------------------------------------------------------------
Public Sub WriteToCell(pValue As Variant, SheetNumber As Long, Column As String, row As Long)
    On Error Resume Next
    XL.Sheets(SheetNumber).Range(Column & row).Value = pValue
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Add a sheet
' Date: November,27 2001 @ 13:33:50
'------------------------------------------------------------
Public Sub AddShet(sNAME As String)
    On Error Resume Next
    XL.Sheets.Add Before:=XL.Sheets(1)
    XL.Sheets(1).Name = sNAME
End Sub



