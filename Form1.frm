VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00D3CAC0&
   Caption         =   "Viplay"
   ClientHeight    =   3570
   ClientLeft      =   6735
   ClientTop       =   4095
   ClientWidth     =   4320
   FillColor       =   &H000020FF&
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4320
   Begin MSComDlg.CommonDialog cmdFile 
      Left            =   2520
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FilterIndex     =   1000
      Flags           =   4100
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      InitDir         =   "10\10"
      Max             =   10
      ToPage          =   10
      Orientation     =   2
   End
   Begin VB.CommandButton cmdSize 
      BackColor       =   &H0036FE4A&
      Caption         =   "up"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      TabIndex        =   3
      Top             =   3360
      Width           =   375
   End
   Begin VB.Timer tmrLeght 
      Interval        =   300
      Left            =   1680
      Top             =   3240
   End
   Begin VB.ListBox lstPL 
      BackColor       =   &H00C2B5A9&
      ForeColor       =   &H00004000&
      Height          =   2400
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   582
      _Version        =   393216
      RecordMode      =   0
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      BackEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      EjectVisible    =   0   'False
      OLEDropMode     =   1
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Line lnPos 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   240
   End
   Begin VB.Shape shpPos 
      BorderColor     =   &H00FF00FF&
      FillColor       =   &H00FF80FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   120
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000020FF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3360
      Width           =   105
   End
   Begin VB.Menu Fail 
      Caption         =   "Файл"
      Begin VB.Menu New 
         Caption         =   "Новый Play List"
         Shortcut        =   ^R
      End
      Begin VB.Menu a1 
         Caption         =   "-"
      End
      Begin VB.Menu OpenFiles 
         Caption         =   "Открыть"
         Shortcut        =   ^F
      End
      Begin VB.Menu a2 
         Caption         =   "-"
      End
      Begin VB.Menu SaveP 
         Caption         =   "Сохранить плейлист"
         Shortcut        =   ^P
      End
      Begin VB.Menu OpenP 
         Caption         =   "Открыть плейлист"
         Shortcut        =   ^O
      End
      Begin VB.Menu a3 
         Caption         =   "-"
      End
      Begin VB.Menu NewTrack 
         Caption         =   "Новый трек"
         Shortcut        =   ^N
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Сохранить трек как"
         Shortcut        =   ^S
      End
      Begin VB.Menu a4 
         Caption         =   "-"
      End
      Begin VB.Menu Delete 
         Caption         =   "Удалить"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Правка"
      Begin VB.Menu AutoNext 
         Caption         =   "Авто прокрутка"
         Checked         =   -1  'True
         Shortcut        =   ^A
      End
      Begin VB.Menu Timer 
         Caption         =   "Таймер"
         Shortcut        =   ^T
      End
      Begin VB.Menu Rndz 
         Caption         =   "Случайно"
      End
      Begin VB.Menu RecMode 
         Caption         =   "Метод записи"
         Begin VB.Menu OverWritten 
            Caption         =   "Перезапись"
         End
         Begin VB.Menu Insert 
            Caption         =   "Вставка"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnWork As Boolean
Dim blnDrag As Boolean
Dim intDragIndex As Integer
Dim intCount As Integer
Dim intIndex As Integer  'выбранный индекс
Dim Ev As New Events

Private Sub AutoNext_Click()
    AutoNext.Checked = Not (AutoNext.Checked)
End Sub

Private Sub cmdSize_Click()
    On Error Resume Next
    Height = 2150
    lstPL.TopIndex = lstPL.ListIndex
End Sub

Private Sub Delete_Click()
Dim intCount As Integer
    For intCount = lstPL.ListCount - 1 To 0 Step -1
        On Error Resume Next
        If lstPL.Selected(intCount) Then
            Ev.DeleteFile intCount
        End If
    Next intCount
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    lstPL_KeyPress (KeyAscii)
End Sub

Private Sub Form_Load()
   On Error Resume Next
    Ev.OpenP App.Path & "\Default.pls"
    Ev.LoadOp
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Height < 2150 Then Height = 2150: lstPL.TopIndex = lstPL.ListIndex
    If Width < 4400 Then Width = 4400
    cmdSize.Top = Height - 970
    cmdSize.Left = Width - 480
    lblTimer.Top = Height - 1050
    shpPos.Width = Width - 400
    lstPL.Width = Width - 360
    lstPL.Top = 800
    MMControl1.Left = Width \ 2 - 1660
    On Error Resume Next
    lstPL.Height = Height - 1800
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Ev.SaveP App.Path & "\Default.pls"
    Ev.SaveOp
End Sub

Private Sub Insert_Click()
    Insert.Checked = True
    OverWritten.Checked = False
    MMControl1.RecordMode = mciRecordInsert
End Sub

Private Sub lblTimer_Click()
    Timer_Click
End Sub

Private Sub lstPL_Click()
    lstPL.ToolTipText = lstPL.ListIndex + 1
End Sub

Private Sub lstPL_DblClick()
    Ev.Play
End Sub


Private Sub lstPL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ev.Play
    End If
    If KeyAscii = 120 Then
      MMControl1_NextClick 1
    End If
    If KeyAscii = 122 Then
        MMControl1_PrevClick 1
    End If
    If KeyAscii = 99 Then
        MMControl1.Command = "play"
    End If
    If KeyAscii = 118 Then
        MMControl1.Command = "pause"
    End If
    If KeyAscii = 98 Then
        MMControl1.Command = "eject"
    End If
    If KeyAscii = 110 Then
        MMControl1.Command = "step 5"
    End If
    If KeyAscii = 109 Then
        MMControl1.Command = "prev"
        MMControl1.Command = "stop"
    End If
    If KeyAscii = 44 Then
        MMControl1.Command = "record"
    End If
End Sub

Private Sub lstPL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnDrag = True
    intDragIndex = lstPL.ListIndex
End Sub

Private Sub lstPL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnDrag And intDragIndex <> lstPL.ListIndex Then
        Ev.Drag intDragIndex, lstPL.ListIndex
    End If
    lstPL.ToolTipText = Y \ 195 + lstPL.TopIndex + 1
    If Val(lstPL.ToolTipText) > lstPL.ListCount Then lstPL.ToolTipText = ""
End Sub

Private Sub lstPL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnDrag = False
    intDragIndex = -1
End Sub

Private Sub lstPL_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strWork() As String
    For intCount = 1 To Data.Files.Count
        'добавляем в список адресов
        strWork = Split(Data.Files(intCount), "\")
        If InStr(strWork(UBound(strWork)), ".") = 0 Then
            Search Data.Files(intCount) & "\"
        Else
            'добавляем в список
            Ev.FlName(lstPL.ListCount) = Data.Files(intCount)
            lstPL.AddItem strWork(UBound(strWork))
        End If
    Next intCount
End Sub

Private Sub Search(CurrentPath As String)
    Dim strWork() As String
    Dim intN As Integer, IntDirectory As Integer
    Dim strFileName As String, strDirectoryList() As String
    ' сначала перечисляем все обычные файлы в текущем каталоге
    On Error Resume Next
    strFileName = Dir(CurrentPath)
    Do While strFileName <> ""
        Ev.FlName(lstPL.ListCount) = CurrentPath & strFileName
        strWork = Split(strFileName, "\")
        'добавляем в список
        lstPL.AddItem strWork(UBound(strWork))
        strFileName = Dir
    Loop
    
    ' теперь формируем временный список подкаталогов
    On Error Resume Next
    strFileName = Dir(CurrentPath, vbDirectory)
    Do While strFileName <> ""
        ' игнорируем текущий и родительские каталоги,
        ' а также файл подкачки Windows NT
        If strFileName <> "." And strFileName <> ".." And strFileName <> "pagefile.sys" Then
            ' игнорируем все, что отличается от каталогов
            If GetAttr(CurrentPath & strFileName) And vbDirectory Then
                IntDirectory = IntDirectory + 1
                ReDim Preserve strDirectoryList(IntDirectory)
                strDirectoryList(IntDirectory) = CurrentPath & strFileName
            End If
        End If
        strFileName = Dir
        ' обрабатываем другие события
    Loop
    ' рекурсивно обрабатываем каждый каталог
    For intN = 1 To (IntDirectory)
        Search strDirectoryList(intN) & "\"
    Next intN
End Sub

Private Sub MMControl1_OLEDragDrop(Data As MCI.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'открываем wav фаил
    'отключаем проигрывание/запись
    MMControl1.Command = "stop"
    MMControl1.Command = "close"
    MMControl1.FileName = Data.Files(1)
    Ev.Play
End Sub

Private Sub MMControl1_PrevClick(Cancel As Integer)
    On Error Resume Next
    If lstPL.ListIndex <> 0 Then
        lstPL.Selected(lstPL.ListIndex) = False
        lstPL.Selected(lstPL.ListIndex - 1) = True
    End If
    Ev.Play
End Sub

Private Sub MMControl1_NextClick(Cancel As Integer)
    AutoNext.Checked = True
    MMControl1.Command = "next"
End Sub

Private Sub MMControl1_StopClick(Cancel As Integer)
    If MMControl1.CanRecord Then MMControl1.Command = "save"
    MMControl1.Command = "prev"
End Sub

Private Sub New_Click()
    Ev.NewPL
End Sub

Private Sub NewTrack_Click()
Dim strWork() As String
'отключаем проигрывание
MMControl1.Command = "stop"
MMControl1.Command = "close"
    MMControl1.RecordEnabled = True
    cmdFile.Filter = ".wav|*.wav"
    cmdFile.ShowSave
    If cmdFile.FileName <> "" Then
        Ev.FlName(lstPL.ListCount) = cmdFile.FileName
        strWork = Split(cmdFile.FileName, "\")
        lstPL.AddItem Str(lstPL.ListCount + 1) & ". " & strWork(UBound(strWork))
        On Error GoTo FileNotFound
        FileCopy App.Path & "\01.wav", cmdFile.FileName
        lstPL.ListIndex = lstPL.ListCount - 1
        Insert_Click
        Ev.Record
    End If
    Exit Sub
FileNotFound:
    MsgBox App.Path & "\01.wav", vbCritical, "File Not Found"
    Exit Sub
End Sub

Private Sub OpenFiles_Click()
Dim a As Long
    With cmdFile
        .DialogTitle = "Open file(s)"
        .Filter = "All files|*.*"
        .ShowOpen
        If .FileName = "" Then Exit Sub
        'добавляем в список адресов
        Ev.FlName(lstPL.ListCount) = .FileName
        lstPL.AddItem Right(.FileName, Len(.FileName) - InStrRev(.FileName, "\"))
    End With
End Sub

Private Sub OpenP_Click()
    cmdFile.Filter = "(файлы плейлиста).pls|*.pls"
    cmdFile.ShowOpen
    If cmdFile.FileName <> "" Then
        Ev.OpenP cmdFile.FileName
    Else
        MsgBox "Ошибка при открытии файла", vbCritical, "Error"
    End If
End Sub

Private Sub OverWritten_Click()
    OverWritten.Checked = True
    Insert.Checked = False
    MMControl1.RecordMode = mciRecordOverwrite
End Sub

Private Sub Rndz_Click()
    Rndz.Checked = Not (Rndz.Checked)
End Sub

Private Sub SaveAs_Click()
     MMControl1.Command = "Save"
     If MMControl1.Error = 0 Then
        MsgBox "Запись сохранена", vbInformation, "Information"
    Else
         MsgBox "Ошибка записи", vbCritical, "Error"
    End If
End Sub

Private Sub SaveP_Click()
    cmdFile.Filter = "(файлы плейлиста).pls|*.pls"
    cmdFile.ShowSave
    If cmdFile.FileName <> "" Then
        Ev.SaveP cmdFile.FileName
    Else
        MsgBox "Ошибка записи", vbCritical, "Error"
    End If
End Sub

Private Sub Timer_Click()
    Timer.Checked = Not (Timer.Checked)
End Sub

Private Sub tmrLeght_Timer()
    lblTimer = (MMControl1.Position - (Timer.Checked And MMControl1.TrackLength)) \ 1000 & " : " & MMControl1.TrackLength \ 1000 & " сек"
    On Error Resume Next
    lnPos.X1 = 120 + shpPos.Width * (MMControl1.Position / (MMControl1.TrackLength + 1))
    lnPos.X2 = lnPos.X1
    'наличие песен
    If lstPL.ListCount And AutoNext.Checked Then
        ' прокрутка песен
        If MMControl1.Position = MMControl1.TrackLength Then
            'случайно
            If Rndz.Checked Then
                    Randomize Timer
                    lstPL.Selected(lstPL.ListIndex) = False
                    lstPL.Selected(Int((lstPL.ListCount - 1) * Rnd)) = True
                    Ev.Play
            Else
                If (lstPL.ListCount - 1) = lstPL.ListIndex Then
                    'прокрутка списка
                    lstPL.ListIndex = 0
                    lstPL.Selected(lstPL.ListCount - 1) = False
                    lstPL.Selected(0) = True
                    Ev.Play
                Else
                    'следующая песня
                    lstPL.Selected(lstPL.ListIndex) = False
                    lstPL.ListIndex = lstPL.ListIndex + 1
                    lstPL.Selected(lstPL.ListIndex) = True
                    Ev.Play
                End If
            End If
        End If
    ElseIf MMControl1.Position = MMControl1.TrackLength And MMControl1.TrackLength <> 0 Then
        Ev.Play
    End If
End Sub
