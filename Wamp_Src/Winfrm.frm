VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wamp"
   ClientHeight    =   3270
   ClientLeft      =   1020
   ClientTop       =   1965
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   10335
   Begin VB.Frame Frame2 
      Caption         =   "Config"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5640
      TabIndex        =   2
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Apply"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   2880
         Width           =   735
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   2400
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   2640
         Width           =   735
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1320
      Top             =   2640
   End
   Begin MSComctlLib.ImageList IL 
      Left            =   480
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2640
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin MSComctlLib.ListView WampV 
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Menu mOpt 
      Caption         =   "Options"
      Begin VB.Menu mC 
         Caption         =   "Config"
      End
      Begin VB.Menu mCom 
         Caption         =   "Commands"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mName 
         Caption         =   "Set Name"
      End
      Begin VB.Menu Ms1 
         Caption         =   "-"
      End
      Begin VB.Menu mRef 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mOnt 
         Caption         =   "Always On Top"
         Checked         =   -1  'True
      End
      Begin VB.Menu mS2 
         Caption         =   "-"
      End
      Begin VB.Menu mX 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mCl 
      Caption         =   "Clipboard"
      Begin VB.Menu mAuto 
         Caption         =   "Auto Copy to Clipboard"
         Checked         =   -1  'True
      End
      Begin VB.Menu msend 
         Caption         =   "Send To Clipboard"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mPre 
         Caption         =   "Preview"
      End
   End
   Begin VB.Menu msnd2 
      Caption         =   "Send To.."
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mYah 
         Caption         =   "Yahoo Name"
      End
      Begin VB.Menu mVP 
         Caption         =   "VP Client"
      End
      Begin VB.Menu mMSN 
         Caption         =   "MSN Name"
      End
   End
End
Attribute VB_Name = "FormW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim LastTrack As String, Force As Boolean
Dim ShowConfig As Boolean, W As Long, MyName As String
Dim CmdArr(15, 2) As String

Private Sub Command1_Click()
Command1.Enabled = False
Clipboard.Clear
Clipboard.SetText WampFormat(Text2)
End Sub

Private Sub Command2_Click()
mPre_Click
End Sub

Private Sub Form_Load()
'On Error Resume Next

WampV.ColumnHeaders(1).Width = 1700
WampV.ColumnHeaders(2).Width = 3500
LoadReg
LoadCmdArr
StayOnTop mOnt.Checked, Me

Me.Width = Frame1.Width + 90
W = Me.Width


For i = 0 To 9
    List1.AddItem CmdArr(i, 1)
Next i
List1.Selected(3) = True
List1.Selected(4) = True
List1.Selected(7) = True
List2.AddItem CmdArr(10, 1) & ":" & CmdArr(11, 1)
List2.AddItem CmdArr(12, 1) & ":" & CmdArr(13, 1)
List2.AddItem CmdArr(14, 1) & ":" & CmdArr(15, 1)
List2.Selected(2) = True
List
Timer1 = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Msg As Long
Msg = X / Screen.TwipsPerPixelX
Select Case Msg
Case WM_LBUTTONUP: Me.Show
Case WM_RBUTTONUP: PopupMenu mOpt
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
ToSysTray Me
End Sub

Private Sub mAuto_Click()
If mAuto.Checked = False Then mAuto.Checked = True Else mAuto.Checked = False
End Sub

Private Sub mC_Click()
mOpt.Enabled = False
If mC.Checked = False Then
    ShowConfig = True
    Timer2 = True
    mC.Checked = True
Else
    mC.Checked = False
    Timer2 = True
End If
End Sub

Private Sub mCom_Click()
Dim Com As String
  Com = "$me          " & vbTab & "Your Name" & vbCrLf & _
        "%tn           " & vbTab & "Track Number" & vbCrLf & _
        "%lc           " & vbTab & "List Count" & vbCrLf & _
        "$t             " & vbTab & "Title" & vbCrLf & _
        "$a             " & vbTab & "Artist" & vbCrLf & _
        "$s             " & vbTab & "Status" & vbCrLf & _
        "%tgm:%tgs   " & vbTab & "Time Gone Mins : Secs" & vbCrLf & _
        "%ttm:%tts   " & vbTab & "Time Total Mins : Secs" & vbCrLf & _
        "%tlm:%tls      " & vbTab & "Time Left Mins : Secs" & vbCrLf & _
        "%tp            " & vbTab & "Time Percentage" & vbCrLf & _
        "%kb/%kh    " & vbTab & "Bitrate kbps / khz" & vbCrLf & _
        "$c             " & vbTab & "Channels (1 = mono, 2 = Stereo)"
        
FormC.Text1 = Com
FormC.Show
StayOnTop True, FormC
End Sub

Private Sub mName_Click()
MyName = InputBox$("Set Value for $me", "Set Your Name", MyName)
End Sub

Private Sub mOnt_Click()
If mOnt.Checked = False Then mOnt.Checked = True Else mOnt.Checked = False
StayOnTop mOnt.Checked, Me
End Sub

Private Sub mPre_Click()
MsgBox WampFormat(Text2), vbApplicationModal + vbSystemModal + vbInformation
End Sub

Private Sub mRef_Click()
Force = True: List
End Sub

Private Sub mSend_Click()
Clipboard.Clear
Clipboard.SetText WampFormat(Text2)
End Sub

Private Sub mX_Click()
SaveReg
KillTray Me
End
End Sub

Private Sub Text2_Change()
Command1.Enabled = True
End Sub

Private Sub Timer1_Timer()
List
End Sub

Sub List()
If IsRunning = False Then Exit Sub
Dim TimeI As Time_Info, TrInf As Track_Info
TimeI = TimeInfo: TrInf = TrackInfo
'uncomment next line to only update on new tune or when Refreh clicked
'If Force = False And LastTrack = TrInf.Number & TrInf.Artist & TrInf.Title Then Exit Sub

Force = False
WampV.ListItems.Clear

CmdArr(0, 2) = MyName
CmdArr(1, 2) = TrInf.Number
CmdArr(2, 2) = TrInf.ListCount
CmdArr(3, 2) = TrInf.Title
CmdArr(4, 2) = TrInf.Artist
CmdArr(5, 2) = TrInf.Status
CmdArr(6, 2) = TimeI.Percent
CmdArr(7, 2) = TrInf.Bitrate.KBPS
CmdArr(8, 2) = TrInf.Bitrate.KHZ
CmdArr(9, 2) = TrInf.Channels
CmdArr(10, 2) = TimeI.TimeGone.Minutes
CmdArr(11, 2) = TimeI.TimeGone.Seconds
CmdArr(12, 2) = TimeI.TimeLeft.Minutes
CmdArr(13, 2) = TimeI.TimeLeft.Seconds
CmdArr(14, 2) = TimeI.TimeTotal.Minutes
CmdArr(15, 2) = TimeI.TimeTotal.Seconds
ListAdd
If LastTrack <> TrInf.Number & TrInf.Artist & TrInf.Title And mAuto.Checked = True Then
    Clipboard.Clear
    Clipboard.SetText WampFormat(Text2)
End If
LastTrack = TrInf.Number & TrInf.Artist & TrInf.Title
End Sub

Function WampFormat(ByVal Value As String) As String
Dim i As Long
For i = 0 To 9
    List1.Selected(i) = False
Next i
For i = 0 To 2
    List2.Selected(i) = False
Next i
For i = 0 To 15
    If InStr(1, Value, CmdArr(i, 0)) <> 0 Then
        Select Case i
            Case 0 To 9
                List1.Selected(i) = True
            Case 10, 11
                List2.Selected(0) = True
            Case 12, 13
                List2.Selected(1) = True
            Case 14, 15
                List2.Selected(2) = True
        End Select
    End If
    Value = Replace$(Value, CmdArr(i, 0), CmdArr(i, 2))
Next i
WampFormat = Value
End Function

Sub ListAdd()
Dim i As Long
For i = 0 To 9
If List1.Selected(i) = True Then
    WampV.ListItems.Add , , CmdArr(i, 1)
    WampV.ListItems(WampV.ListItems.Count).ListSubItems.Add , , CmdArr(i, 2)
End If
Next i

If List2.Selected(0) = True Then
    WampV.ListItems.Add , , "Time Played"
    WampV.ListItems(WampV.ListItems.Count).ListSubItems.Add , , CmdArr(10, 2) & ":" & CmdArr(11, 2)
End If
If List2.Selected(1) = True Then
    WampV.ListItems.Add , , "Time Remaining"
    WampV.ListItems(WampV.ListItems.Count).ListSubItems.Add , , CmdArr(12, 2) & ":" & CmdArr(13, 2)
End If
If List2.Selected(2) = True Then
    WampV.ListItems.Add , , "Total Length"
    WampV.ListItems(WampV.ListItems.Count).ListSubItems.Add , , CmdArr(14, 2) & ":" & CmdArr(15, 2)
End If

End Sub

Private Sub Timer2_Timer()
If ShowConfig = True Then
Dim i As Long
For i = Me.Width To (Me.Width + Frame2.Width + 100) Step 4
    Me.Width = i
    DoEvents
Next i
ShowConfig = False
Else
For i = Me.Width To W Step -4
    Me.Width = i
    DoEvents
Next i
End If
Timer2 = False
mOpt.Enabled = True
End Sub

Sub LoadCmdArr()
CmdArr(0, 0) = "$me": CmdArr(0, 1) = "User Name": CmdArr(0, 2) = MyName
CmdArr(1, 0) = "%tn": CmdArr(1, 1) = "Track No"
CmdArr(2, 0) = "%lc": CmdArr(2, 1) = "List Count"
CmdArr(3, 0) = "$t": CmdArr(3, 1) = "Track"
CmdArr(4, 0) = "$a": CmdArr(4, 1) = "Artist"
CmdArr(5, 0) = "$s": CmdArr(5, 1) = "Status"
CmdArr(6, 0) = "%tp": CmdArr(6, 1) = "Time Percentage"
CmdArr(7, 0) = "%kb": CmdArr(7, 1) = "KBPS"
CmdArr(8, 0) = "%kh": CmdArr(8, 1) = "KHZ"
CmdArr(9, 0) = "$c": CmdArr(9, 1) = "Channels"
CmdArr(10, 0) = "%tgm": CmdArr(10, 1) = "Gone (M)"
CmdArr(11, 0) = "%tgs": CmdArr(11, 1) = "Gone (S)"
CmdArr(12, 0) = "%tlm": CmdArr(12, 1) = "Left (M)"
CmdArr(13, 0) = "%tls": CmdArr(13, 1) = "Left (S)"
CmdArr(14, 0) = "%ttm": CmdArr(14, 1) = "Total (M)"
CmdArr(15, 0) = "%tts": CmdArr(15, 1) = "Total (S)"
End Sub

Sub LoadReg()
Dim Appname As String, Section As String, i As Long
Appname = App.EXEName
Section = "WAMP Settings"
MyName = GetSetting(Appname, Section, "User", "User")
Text2 = GetSetting(Appname, Section, "WAMP", Text2)
mOnt.Checked = GetSetting(Appname, Section, "On top", mOnt.Checked)
mAuto.Checked = GetSetting(Appname, Section, "Auto Clip", mAuto.Checked)
End Sub

Sub SaveReg()
Dim Appname As String, Section As String, i As Long
Appname = App.EXEName
Section = "WAMP Settings"
SaveSetting Appname, Section, "User", MyName
SaveSetting Appname, Section, "WAMP", Text2
SaveSetting Appname, Section, "On top", mOnt.Checked
SaveSetting Appname, Section, "Auto Clip", mAuto.Checked
End Sub
