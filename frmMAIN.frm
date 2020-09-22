VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excel Class Demo"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEXPORT 
      Caption         =   "&Export"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgSMALL 
      Left            =   2640
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.TextBox txtNAME 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "Demo Export"
      Top             =   120
      Width           =   4815
   End
   Begin MSComctlLib.ListView lvLIST 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblLINK 
      Caption         =   "http://vbasic.iscool.net"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblLABEL 
      Caption         =   "Sheet Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Get File name and on success of a name, export
'                to that file.
' Date: November,27 2001 @ 13:31:10
'------------------------------------------------------------
Private Sub cmdEXPORT_Click()
    On Error Resume Next
    Dim fNAME As String, obj As CDLG
    Set obj = New CDLG
    fNAME = ""
    obj.VBGetSaveFileName fNAME, , , "Excel Workbook (*.xls)|*.xls", , CurDir, "Export to:", "*.xls"
    If fNAME <> "" Then
        Screen.MousePointer = vbHourglass
        ExportIt fNAME
        Screen.MousePointer = vbDefault
        MsgBox "Export to:" & fNAME & " complete.", vbInformation, "Done"
    End If
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Exports the ListView to the selected file using
'                the CEXCEL class.
' Date: November,27 2001 @ 11:38:15
'------------------------------------------------------------
Private Sub ExportIt(fNAME As String)
    On Error GoTo ErrorExportIt
    Dim obj As CEXCEL, itm As ListItem, x As Long
    '------------------------------------------------------------
    ' Extract that template workbook file from the
    ' resource file to the selected file name
    '------------------------------------------------------------
    GenFileFromRes 101, "XLS", "XLS", , , fNAME
    Set obj = New CEXCEL
    With obj
        '------------------------------------------------------------
        ' Opens the workbook hidden.
        '------------------------------------------------------------
        .OpenExcel fNAME
        '------------------------------------------------------------
        ' Set the name of Sheet 1 to the text entered.
        '------------------------------------------------------------
        If Me.txtNAME.Text <> "" Then .XL.Sheets(1).Name = Me.txtNAME.Text
        '------------------------------------------------------------
        ' Loop the column headers and write them to row 1 on sheet 1
        '------------------------------------------------------------
        For x = 1 To lvLIST.ColumnHeaders.Count
            .WriteToCell lvLIST.ColumnHeaders(x).Text, 1, Chr(64 + x), 1
        Next x
        '------------------------------------------------------------
        ' Loop each listitem and write out the data
        '------------------------------------------------------------
        For x = 1 To lvLIST.ListItems.Count
            Set itm = lvLIST.ListItems(x)
            .WriteToCell itm.Text, 1, "A", x + 1
            .WriteToCell itm.SubItems(1), 1, "B", x + 1
            .WriteToCell itm.SubItems(2), 1, "C", x + 1
            .WriteToCell itm.SubItems(3), 1, "D", x + 1
            .WriteToCell itm.SubItems(4), 1, "E", x + 1
        Next x
        '------------------------------------------------------------
        ' I wanted to show the XL property of the class.
        '  It is a direct link back to the Excel Object
        ' pointing to the workbook.  If you know Excel
        ' VBA you can type in any code that Excel will
        ' understand using it.  I resized the first five
        ' columns to 11 with this code then set the selected
        ' cell back to A1
        '------------------------------------------------------------
        .XL.Sheets(1).Columns("A:E").ColumnWidth = 11
        .XL.Sheets(1).Range("A1").Select
        '------------------------------------------------------------
        ' Here I am adjusting the Page Setup [you can see
        ' it when you hit print preview on the file in
        ' Excel.
        '------------------------------------------------------------
        .XL.Sheets(1).PageSetUp.CenterHeader = "DEMO EXPORT"
        .XL.Sheets(1).PageSetUp.LeftFooter = "http://vbasic.iscool.net"
        .XL.Sheets(1).PageSetUp.RightFooter = "Clint LaFever"
        '------------------------------------------------------------
        ' Create a pivot table from the data
        '------------------------------------------------------------
        .XL.Sheets(1).PivotTableWizard SourceType:=1, SourceData:="'" & .XL.Sheets(1).Name _
            & "'!R1C1:R" & lvLIST.ListItems.Count + 1 & "C5", TableDestination:="", TableName:="PivotTable1"
        .XL.Sheets(1).PivotTables("PivotTable1").AddFields RowFields:=Array("ITEM", "Data"), _
            ColumnFields:="STORE"
        With .XL.Sheets(1).PivotTables("PivotTable1").PivotFields(DateAdd("d", -1, Date))
            .Orientation = 4
            .Position = 1
        End With
        With .XL.Sheets(1).PivotTables("PivotTable1").PivotFields(DateAdd("d", -2, Date))
            .Orientation = 4
            .Position = 2
        End With
        .XL.Sheets(1).PivotTables("PivotTable1").PivotFields(DateAdd("d", -3, Date)).Orientation = 4
        '------------------------------------------------------------
        ' Here I am adjusting the Page Setup [you can see
        ' it when you hit print preview on the file in
        ' Excel.
        ' Note: When I created the Pivot Table, it created a new sheet 1 and moved
        ' the other to sheet 2
        '------------------------------------------------------------
        .XL.Sheets(1).PageSetUp.CenterHeader = "DEMO EXPORT"
        .XL.Sheets(1).PageSetUp.LeftFooter = "http://vbasic.iscool.net"
        .XL.Sheets(1).PageSetUp.RightFooter = "Clint LaFever"
        .XL.Sheets(1).Name = "Pivot"
        .CloseExcel True
    End With
    Set obj = Nothing
    Exit Sub
ErrorExportIt:
    If Not obj Is Nothing Then
        obj.CloseExcel True
        Set obj = Nothing
    End If
    MsgBox Err & ":Error in ExportIt.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub Form_Load()
    On Error Resume Next
    SetFormIcon Me, iconDEFAULT
    AddImage imgSMALL, iconDEFAULT, IMG_SIXTEEN
    InitlvLIST
    FilllvLIST "STORE A"
    FilllvLIST "STORE B"
    FilllvLIST "STORE C"
    FilllvLIST "STORE D"
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Initialize the ListView
' Date: November,27 2001 @ 11:29:19
'------------------------------------------------------------
Private Sub InitlvLIST()
    On Error GoTo ErrorInitlvLIST
    With lvLIST
        .SmallIcons = imgSMALL
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .GridLines = True
        .ColumnHeaders.Add , , "ITEM"
        .ColumnHeaders.Add , , "STORE"
        .ColumnHeaders.Add , , DateAdd("d", -1, Date)
        .ColumnHeaders.Add , , DateAdd("d", -2, Date)
        .ColumnHeaders.Add , , DateAdd("d", -3, Date)
    End With
    Exit Sub
ErrorInitlvLIST:
    MsgBox Err & ":Error in InitlvLIST.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Fill the ListView with dummy data
' Date: November,27 2001 @ 11:30:57
'------------------------------------------------------------
Private Sub FilllvLIST(StoreName As String)
    On Error GoTo ErrorFilllvLIST
    Dim x As Long, itm As ListItem
    Randomize
    For x = 1 To 15
        Set itm = lvLIST.ListItems.Add(, , "ITEM " & x, , 1)
        itm.SubItems(1) = StoreName
        itm.SubItems(2) = Int((50 * Rnd) + 1)
        itm.SubItems(3) = Int((45 * Rnd) + 1)
        itm.SubItems(4) = Int((75 * Rnd) + 1)
    Next x
    Exit Sub
ErrorFilllvLIST:
    MsgBox Err & ":Error in FilllvLIST.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub

Private Sub lblLINK_Click()
    On Error Resume Next
    ExecuteLink lblLINK.Caption
End Sub
