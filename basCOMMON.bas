Attribute VB_Name = "basCOMMON"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Public Sub ExecuteLink(LINK As String)
    On Error Resume Next
    Dim lRet As Long
    If LINK <> "" Then
        lRet = ShellExecute(0, "open", LINK, "", App.Path, SW_SHOWNORMAL)
        If lRet >= 0 And lRet <= 32 Then
            MsgBox "Error jumping to:" & LINK, 48, "Warning"
        End If
    End If
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever - [lafeverc@saic.com]
' Purpose:  Extracts a file from the custom resource file
'                to the local hard drive.
' Parameters:  resID=ID of resource  :  resSECTION=Section of custom resource ie. CUSTOM
'                     fEXT=Extension for new file  :  fPATH=Destination path, default is App.Path
'                     fNAME=Name for new file, default is TEMP
' Returns:  Full path and file name of file created
' Example:  retSTR=GenFileFromRes(101,"CUSTOM","JPG",,"IMAGE")
' Date: December,17 1999 @ 10:50:58
'------------------------------------------------------------
Public Function GenFileFromRes(resID As Long, resSECTION As String, fEXT As String, Optional fPath As String = "", Optional fNAME As String = "temp", Optional FullName As String = "") As String
    On Error GoTo ErrorGenFileFromRes
    Dim resBYTE() As Byte
    If fPath = "" Then fPath = App.Path
    If fNAME = "" Then fNAME = "temp"
    '------------------------------------------------------------
    ' Get the file out of the resource file
    '------------------------------------------------------------
    resBYTE = LoadResData(resID, resSECTION)
    '------------------------------------------------------------
    ' Open destination
    '------------------------------------------------------------
    If FullName = "" Then
        Open fPath & "\" & fNAME & "." & fEXT For Binary Access Write As #1
    Else
        Open FullName For Binary Access Write As #1
    End If
    '------------------------------------------------------------
    ' Write it out
    '------------------------------------------------------------
    Put #1, , resBYTE
    '------------------------------------------------------------
    ' Close it
    '------------------------------------------------------------
    Close #1
    If FullName = "" Then
        GenFileFromRes = fPath & "\" & fNAME & "." & fEXT
    Else
        GenFileFromRes = FullName
    End If
    Exit Function
ErrorGenFileFromRes:
    GenFileFromRes = ""
    MsgBox Err & ":Error in GenFileFromRes.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@usa.net]
' Purpose:  Used to Add an image to a ImageList from the resource file.  Note.  AppIcons must be declared.
' Parameters:
' Example:
' Date: July,21 1998 @ 18:22:18
'------------------------------------------------------------
Public Sub AddImage(imgLIST As ImageList, resICONVAL As AppIcons, Optional imgSIZE As IMG_SIZE = IMG_ALREADYSET, Optional CustomHeight As Long = 16, Optional CustomWidth As Long = 16)
    On Error Resume Next
    With imgLIST
        If imgSIZE <> IMG_ALREADYSET Then
            If imgSIZE <> IMG_CUSTOM Then
                .ImageHeight = imgSIZE
                .ImageWidth = imgSIZE
            Else
                .ImageHeight = CustomHeight
                .ImageWidth = CustomWidth
            End If
        End If
        .ListImages.Add , , LoadResPicture(resICONVAL, vbResIcon)
    End With
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@usa.net]
' Purpose:  Changes the size of icons within an ImageList at RunTime.
' Parameters:
' Example:
' Date: July,21 1998 @ 18:22:47
'------------------------------------------------------------
Public Sub ChangeImageSize(imgLIST As ImageList, imgSIZE As IMG_SIZE, Optional CustomHeight As Long = 16, Optional CustomWidth As Long = 16)
    On Error Resume Next
    With imgLIST
        If imgSIZE <> IMG_ALREADYSET Then
            If imgSIZE <> IMG_CUSTOM Then
                .ImageHeight = imgSIZE
                .ImageWidth = imgSIZE
            Else
                .ImageHeight = CustomHeight
                .ImageWidth = CustomHeight
            End If
        End If
    End With
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@usa.net]
' Purpose:  Used to set a given form's Icon property to an icon from the Resource File.  Note the use of AppIcons
' Parameters:
' Example:
' Date: July,21 1998 @ 19:25:18
'------------------------------------------------------------
Public Sub SetFormIcon(frm As Form, lngICON As AppIcons)
    On Error Resume Next
    frm.Icon = LoadResPicture(lngICON, vbResIcon)
End Sub


