Public user As String
Public foldername2 As String
Public folderPath As String
Public versioNumber As String
Public ws As Worksheet
Public path As String
Public Filename As String

'Beging of prompt library
#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" _
        () As Long
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
        (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" _
        (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" _
        (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" _
        (ByVal hHook As LongPtr) As Long
    Private hHook As LongPtr        ' handle to the Hook procedure (global variable)
#Else
    Private Declare Function GetCurrentThreadId Lib "kernel32" _
        () As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
        (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" _
        (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
    Private Declare Function CallNextHookEx Lib "user32" _
        (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" _
        (ByVal hHook As Long) As Long
    Private hHook As Long           ' handle to the Hook procedure (global variable)
#End If
' Hook flags (Computer Based Training)
Private Const WH_CBT = 5            ' hook type
Private Const HCBT_ACTIVATE = 5     ' activate window
' MsgBox constants (these are enumerated by VBA)
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7 (these are button IDs)
' for 1 button, use vbOKOnly = 0 (OK button with ID vbOK returned)
' for 2 buttons, use vbOKCancel = 1 (vbOK, vbCancel) or vbYesNo = 4 (vbYes, vbNo) or vbRetryCancel = 5 (vbRetry, vbCancel)
' for 3 buttons, use vbAbortRetryIgnore = 2 (vbAbort, vbRetry, vbIgnore) or vbYesNoCancel = 3 (vbYes, vbNo, vbCancel)
' Module level global variables
Private sMsgBoxDefaultLabel(1 To 7) As String
Private sMsgBoxCustomLabel(1 To 7) As String
Private bMsgBoxCustomInit As Boolean

Private Sub MsgBoxCustom_Init()
' Initialize default button labels for Public Sub MsgBoxCustom
    Dim nID As Integer
    Dim vA As Variant               ' base 0 array populated by Array function (must be Variant)
    vA = VBA.Array(vbNullString, "OK", "Cancel", "Abort", "Retry", "Ignore", "Yes", "No")
    For nID = 1 To 7
        sMsgBoxDefaultLabel(nID) = vA(nID)
        sMsgBoxCustomLabel(nID) = sMsgBoxDefaultLabel(nID)
    Next nID
    bMsgBoxCustomInit = True
End Sub

Public Sub MsgBoxCustom_Set(ByVal nID As Integer, Optional ByVal vLabel As Variant)
' Set button nID label to CStr(vLabel) for Public Sub MsgBoxCustom
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7
' If nID is zero, all button labels will be set to default
' If vLabel is missing, button nID label will be set to default
' vLabel should not have more than 10 characters (approximately)
    If nID = 0 Then Call MsgBoxCustom_Init
    If nID < 1 Or nID > 7 Then Exit Sub
    If Not bMsgBoxCustomInit Then Call MsgBoxCustom_Init
    If IsMissing(vLabel) Then
        sMsgBoxCustomLabel(nID) = sMsgBoxDefaultLabel(nID)
    Else
        sMsgBoxCustomLabel(nID) = CStr(vLabel)
    End If
End Sub

Public Sub MsgBoxCustom_Reset(ByVal nID As Integer)
' Reset button nID to default label for Public Sub MsgBoxCustom
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7
' If nID is zero, all button labels will be set to default
    Call MsgBoxCustom_Set(nID)
End Sub

#If VBA7 Then
    Private Function MsgBoxCustom_Proc(ByVal lMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Function MsgBoxCustom_Proc(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
' Hook callback function for Public Function MsgBoxCustom
    Dim nID As Integer
    If lMsg = HCBT_ACTIVATE And bMsgBoxCustomInit Then
        For nID = 1 To 7
            SetDlgItemText wParam, nID, sMsgBoxCustomLabel(nID)
        Next nID
    End If
    MsgBoxCustom_Proc = CallNextHookEx(hHook, lMsg, wParam, lParam)
End Function

Public Sub MsgBoxCustom( _
    ByRef vID As Variant, _
    ByVal sPrompt As String, _
    Optional ByVal vButtons As Variant = 0, _
    Optional ByVal vTitle As Variant, _
    Optional ByVal vHelpfile As Variant, _
    Optional ByVal vContext As Variant = 0)
' Display standard VBA MsgBox with custom button labels
' Return vID as result from MsgBox corresponding to clicked button (ByRef...Variant is compatible with any type)
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7
' Arguments sPrompt, vButtons, vTitle, vHelpfile, and vContext match arguments of standard VBA MsgBox function
' This is Public Sub instead of Public Function so it will not be listed as a user-defined function (UDF)
    hHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxCustom_Proc, 0, GetCurrentThreadId)
    If IsMissing(vHelpfile) And IsMissing(vTitle) Then
        vID = MsgBox(sPrompt, vButtons)
    ElseIf IsMissing(vHelpfile) Then
        vID = MsgBox(sPrompt, vButtons, vTitle)
    ElseIf IsMissing(vTitle) Then
        vID = MsgBox(sPrompt, vButtons, , vHelpfile, vContext)
    Else
        vID = MsgBox(sPrompt, vButtons, vTitle, vHelpfile, vContext)
    End If
    If hHook <> 0 Then UnhookWindowsHookEx hHook
End Sub
'End of prompt library

Sub filename_cellvalue()
Dim cellvalue As String
user = Environ("USERNAME ")
Range("N10").NumberFormat = "@"
Filename = Range("N10")
line = Range("T12")

If Not Workbooks("Kollinpiippaus.xlsm").Worksheets("Sheet1").Range("G11") = "" Then
    Workbooks("Kollinpiippaus.xlsm").Worksheets("Sheet1").Range("G8") = "Päällä"
Else
    Workbooks("Kollinpiippaus.xlsm").Worksheets("Sheet1").Range("G8") = "Ei päällä"
End If

CreateFolders "Line 1"
CreateFolders "Line 2"
CreateFolders "Line 3"

If Filename <> "" Then
Call Timer

Range("A47") = Day(Now) & "." & Month(Now) & "." & Year(Now)

'Checking if all the batteries are read
Set ws = Worksheets("Sheet1")
Dim bIsEmpty As Boolean
bIsEmpty = False
If ws.Application.WorksheetFunction.CountBlank(ws.Range("N2:N9")) > 0 Then
    bIsEmpty = True
End If


'If the user has not read all the batteries
If bIsEmpty = True Then
    MsgBoxCustom_Set vbYes, "Tallenna"
    MsgBoxCustom_Set vbNo, "Älä tallenna"
    MsgBoxCustom ans, "Et ole lukenut kaikkia akkuja. Tallenna silti?", (vbYesNo + vbQuestion)
    If ans = 6 Then
        bIsEmpty = False
    End If
    If ans = 7 Then
        Range("N10").Clear
        ActiveSheet.Range("N10").Select
        ActiveSheet.Range("N10").NumberFormat = "@"
        'ActiveSheet.Range("T13") = ""
        'ActiveSheet.Range("T11") = ""
    End If
End If


'If the user scans a wrong barcode
Dim length As Integer
length = Len(Range("N10"))
If length <> 28 And bIsEmpty = False Then
    MsgBox Prompt:="Luit väärän kollinumeron. Lue uudestaan.", Buttons:=vbExclamation
    Range("N10").Clear
    ActiveSheet.Range("N10").Select
    ActiveSheet.Range("N10").NumberFormat = "@"
End If


'Checking if the user has read some battery twice or more by mistake
Dim exists As Integer
For j = 2 To 9
  Dim toCheck As String
  toCheck = Range("N" & j)
  For k = 2 To 9
    If toCheck = Range("N" & k) Then
      exists = exists + 1
    End If
  Next k
Next j
If exists > 8 And ans < 6 And ans > 7 Then
  MsgBox Prompt:="Luit jonkin akun useampaan kertaan.", Buttons:=vbExclamation
  Range("N10").Clear
  ActiveSheet.Range("N10").NumberFormat = "@"
  Range("N10").Select
  bIsEmpty = True
End If


'Checking which line it is
Dim s As String
s = Range("N3")
Dim lineNumber As String
lineNumber = Mid(s, 22, 1)


Dim lastrow As Long
Dim i As Integer
Dim s2 As String
s2 = ActiveSheet.Range("N3")

Dim batteryversio As String
Dim variantti As String
versioNumber = Mid(s2, 2, 10)

'Checking that versiot-list has a matching battery version so we can be sure that dynamically read versioNumber is right
lastrow = Workbooks("Kollinpiippaus.xlsm").Worksheets("versiot").Range("B30000").End(xlUp).Row
For i = 1 To lastrow
    If InStr(1, LCase(Workbooks("Kollinpiippaus.xlsm").Worksheets("versiot").Range("B" & i)), versioNumber) <> 0 Then
      batteryversio = (Workbooks("Kollinpiippaus.xlsm").Worksheets("versiot").Range("B" & i))
      'MsgBox ("batteryversio: " & batteryversio)
      End If
Next i

'If a matching battery was not found in the list
If batteryversio = "" Then
  versioNumber = "null"
  MsgBox Prompt:="Akun versionumeroa ei löytynyt. Versioiden tulee olla Kollinpiippausmakron versiot-välilehdellä.", Buttons:=vbExclamation
  Range("N10").Clear
End If

folderPath = Application.Workbooks("Kollinpiippaus.xlsm").path

If length = 28 And bIsEmpty = False And lineNumber = "3" And versioNumber <> "null" Then
  SaveFile "Line 3", 1300
End If

If length = 28 And bIsEmpty = False And lineNumber = "2" And versioNumber <> "null" Then
  SaveFile "Line 2", 1200
End If

If length = 28 And bIsEmpty = False And lineNumber = "1" And versioNumber <> "null" Then
  SaveFile "Line 1", 1100
End If

If lineNumber <> "3" And lineNumber <> "2" And lineNumber <> "1" And bIsEmpty = False Then
  MsgBox Prompt:="Linjanumeroa ei löytynyt. Luitko akut jo yllä oleviin kenttiin?", Buttons:=vbExclamation
    Range("N10").Clear
End If

Else:
Call Timer

End If
End Sub

Sub Timer()
If Workbooks("Kollinpiippaus.xlsm").Worksheets("sheet1").Range("G11") <> "" Then
Application.OnTime Now + TimeValue("00:00:03"), "filename_cellvalue"
End If
End Sub

Sub Stopmacro()
Application.ActiveWorkbook.Worksheets("Sheet1").Range("g11").Value = ""
Workbooks("Kollinpiippaus.xlsm").Worksheets("Sheet1").Range("G8") = "Ei päällä"
End Sub

Sub PictureKiller()
    Dim asd As Shape, rng As Range
    Set rng = Range("A1:L7")
    For Each asd In ActiveSheet.Shapes
        If Intersect(rng, asd.TopLeftCell) Is Nothing Then
        Else
            asd.Delete
        End If
    Next asd
End Sub

Function CreateFolders(line As String)
Dim foldername As String
Dim folderexists As String
foldername = "C:\Users\" & user & "\Desktop\" & line & "\" & MonthName(Month(Now)) & " " & Year(Now) & "\"
folderexists = Dir(foldername, vbDirectory)
If folderexists = "" Then
    VBA.FileSystem.MkDir (foldername)
End If
Dim folderexists2 As String
foldername2 = "C:\Users\" & user & "\Desktop\" & line & "\" & MonthName(Month(Now)) & " " & Year(Now) & "\" & Day(Now) & "." & Month(Now) & "." & Year(Now)
folderexists2 = Dir(foldername2, vbDirectory)
If folderexists2 = "" Then
    VBA.FileSystem.MkDir (foldername2)
End If
path = "C:\Users\" & user & "\Desktop\"
End Function

Function SaveFile(line As String, l As Integer)
Dim PicPath1 As String, Pic1 As Picture, ImageCell1 As Range
PicPath1 = folderPath & "\versiot\" & "A" & versioNumber & ".emf"
ActiveSheet.Range("A1").Select
Set ImageCell1 = ActiveCell.MergeArea
Dim r1 As Range
Set r1 = ws.Range("A1:L54")
Dim sht As Worksheet: Set sht = ActiveSheet
With sht.Shapes
  .AddPicture _
  Filename:=PicPath1, _
  LinkToFile:=msoFalse, _
  SaveWithDocument:=msoTrue, _
  Left:=r1.Left, _
  Top:=r1.Top, _
  Width:=r1.Width, _
  Height:=r1.Height
End With
Range("N10").Clear
Range("T12") = l
ActiveWorkbook.SaveAs Filename:=path & line & "\" & MonthName(Month(Now)) & " " & Year(Now) & "\" & Day(Now) & "." & Month(Now) & "." & Year(Now) & "\" & Filename & ".xlsx"
Range("T12") = ""
Range("N2:N10") = ""
Range("N2").Select
Call PictureKiller
Beep
End Function




