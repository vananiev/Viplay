VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strFiles() As String

Public Sub Play()
    'открываем wav фаил
    'отключаем проигрывание/запись
    With frmMain
        If .lstPL.ListIndex = -1 Then Exit Sub
        .MMControl1.Command = "stop"
        .MMControl1.Command = "close"
        .MMControl1.FileName = strFiles(.lstPL.ListIndex)
        'открываем контрол
        On Error Resume Next
        .MMControl1.Command = "open"
         'запускаем запись и проверяем открываемость записи
        .MMControl1.Command = "play"
        If .MMControl1.Error = 257 Then
            ' неповторяемая запись "[can`t play]"
            If Right(.lstPL.List(.lstPL.ListIndex), 12) <> "[can`t play]" Then
                .lstPL.List(.lstPL.ListIndex) = .lstPL.List(.lstPL.ListIndex) & " [can`t play]"
            End If
        End If
        .Caption = Str(.lstPL.ListIndex + 1) & ". " & .lstPL.List(.lstPL.ListIndex) & " - Viplay"
    End With
End Sub

Public Sub Record()
    'открываем wav фаил
    'отключаем проигрывание/запись
    With frmMain
        .MMControl1.Command = "stop"
        .MMControl1.Command = "close"
        .MMControl1.FileName = strFiles(.lstPL.ListIndex)
        'открываем контрол
        .MMControl1.Command = "open"
         'запускаем запись
        .MMControl1.Command = "record"
        .Caption = .lstPL.List(.lstPL.ListIndex)
    End With
End Sub

Public Property Get FlName(ByVal Index As Integer) As String
    If Index < 0 Or Index > 10000 Then MsgBox "Mistake of namber sing", vbCritical, "Error": Index = 10000
    FlName = strFiles(Index)
End Property

Public Property Let FlName(ByVal Index As Integer, ByVal NewName As String)
    If Index < 0 Or Index > UBound(strFiles) Then Index = UBound(strFiles) + 1: ReDim Preserve strFiles(Index)
    strFiles(Index) = NewName
End Property

Public Sub NewPL()
Dim intCnt As Integer
    ReDim strFiles(0)
    frmMain.lstPL.Clear
    frmMain.lstPL.ListIndex = -1
End Sub

Public Sub DeleteFile(ByVal Index As Integer)
Dim intCnt As Integer
If Index < 0 Then Exit Sub
    For intCnt = Index To UBound(strFiles) - 1
        strFiles(intCnt) = strFiles(intCnt + 1)
    Next intCnt
    frmMain.lstPL.RemoveItem (Index)
End Sub

Public Sub Drag(From As Integer, Dist As Integer)
Dim strWork As String
Dim strWorkMem As String
Dim intW As Integer
    ' перетаскивание на PL
    With frmMain
        strWork = .lstPL.List(From)
        strWorkMem = FlName(From)
        For intW = From To Dist + Sgn(From - Dist) Step Sgn(Dist - From)
            .lstPL.List(intW) = .lstPL.List(intW + Sgn(Dist - From))
            'перетаскивание в памяти
            FlName(intW) = FlName(intW + Sgn(Dist - From))
        Next intW
        FlName(Dist) = strWorkMem
        .lstPL.List(Dist) = strWork
        From = Dist
    End With
End Sub

Public Sub SaveP(FileName As String)
Dim intCount As Integer
    Open FileName For Output As #1
        Do While intCount <> frmMain.lstPL.ListCount
            Print #1, strFiles(intCount)
            intCount = intCount + 1
        Loop
    Close #1
End Sub

Public Sub OpenP(FileName As String)
Dim intCount As Integer
Dim a As String
    NewPL
    On Error Resume Next
    Open FileName For Input As #1
        Do While Not (EOF(1))
            ReDim Preserve strFiles(intCount)
            Line Input #1, strFiles(intCount)
            frmMain.lstPL.AddItem Right(strFiles(intCount), Len(strFiles(intCount)) - InStrRev(strFiles(intCount), "\"))
            a = Right(strFiles(intCount), Len(strFiles(intCount)) - InStrRev(strFiles(intCount), "\"))
            intCount = intCount + 1
        Loop
    Close #1
    Exit Sub
ErrorOpen:
    MsgBox "Ошибка при открытии файла", vbCritical, "Error"
    Close #1
End Sub

Public Sub SaveOp()
    With frmMain
        SaveSetting App.Title, "Options", "AutoNext", CStr(.AutoNext.Checked)
        SaveSetting App.Title, "Options", "Insert", CStr(.Insert.Checked)
        SaveSetting App.Title, "Options", "OverWritten", CStr(.OverWritten.Checked)
        SaveSetting App.Title, "Options", "Timer", CStr(.Timer.Checked)
        SaveSetting App.Title, "Options", "Rnd", .Rndz.Checked
        SaveSetting App.Title, "Options", "Width", .Width
        SaveSetting App.Title, "Options", "Height", .Height
        SaveSetting App.Title, "Options", "Top", Str(.Top)
        SaveSetting App.Title, "Options", "Left", Str(.Left)
        SaveSetting App.Title, "Options", "Selected", .lstPL.ListIndex
    End With
End Sub

Public Sub LoadOp()
    With frmMain
        .AutoNext.Checked = GetSetting(App.Title, "Options", "AutoNext", True)
        .Insert.Checked = GetSetting(App.Title, "Options", "Insert", True)
        .OverWritten.Checked = GetSetting(App.Title, "Options", "OverWritten", False)
        .Timer.Checked = GetSetting(App.Title, "Options", "Timer", False)
        .Rndz.Checked = GetSetting(App.Title, "Options", "Rnd", False)
        .Width = GetSetting(App.Title, "Options", "Width", 4440)
        .Height = GetSetting(App.Title, "Options", "Height", 4380)
        .lstPL.ListIndex = GetSetting(App.Title, "Options", "Selected", 0) - (1 And .AutoNext.Checked)
        If .lstPL.ListIndex < -1 Then .lstPL.ListIndex = -1
        If .lstPL.ListIndex <> -1 Then .lstPL.Selected(.lstPL.ListIndex) = Not (.AutoNext.Checked)
        On Error GoTo Err
        .Top = GetSetting(App.Title, "Options", "Top", Str(Screen.Height / 2 - 2000))
        On Error GoTo Err
        .Left = GetSetting(App.Title, "Options", "Left", Str(Screen.Width / 2 - 2500))
    End With
    Exit Sub
Err:
    frmMain.Top = 0
    frmMain.Left = 0
End Sub
