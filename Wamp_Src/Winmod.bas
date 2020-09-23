Attribute VB_Name = "WinMod"

Option Explicit
'for winamp 1,2,4,5
'winamp 3 used a different engine. rather like Halloween 3, that one about the masks? weird
'Based on a Winamp Class downloaded from PSC

Private hwndWinamp As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Const WM_USER = &H400

Private Enum WM_USER_MSGS
    wmVersion = 0
    wmPlaybackStatus = 104
    wmTrackTimeInfo = 105
    wmPlaylistLength = 124
    wmTrackRateInfo = 126
End Enum

Private Type Track_Time
    Minutes     As Long
    Seconds     As Long
End Type

Public Type Time_Info
    TimeGone    As Track_Time
    TimeLeft    As Track_Time
    TimeTotal   As Track_Time
    Percent     As Long
End Type

Private Type Bit_Rate
    KHZ         As Long
    KBPS        As Long
End Type

Public Type Track_Info
    Title       As String
    Artist      As String
    Number      As Long
    Bitrate     As Bit_Rate
    Channels    As Long
    Status      As String
    ListCount   As Long
End Type

Public Property Get IsRunning() As Boolean
hwndWinamp = FindWindow("Winamp v1.x", vbNullString)
Select Case hwndWinamp
Case 0: IsRunning = False
Case Else: IsRunning = True
End Select
End Property

Public Property Get WinampVersion() As String
WinampVersion = SendUser(wmVersion)
End Property

Private Function SendUser(ID As WM_USER_MSGS, Optional Data As Long) As Long
hwndWinamp = FindWindow("Winamp v1.x", vbNullString)
SendUser = SendMessage(hwndWinamp, WM_USER, Data, ID)
End Function

Public Property Get TrackInfo() As Track_Info
TrackInfo = GetInfo
End Property

Public Property Get TimeInfo() As Time_Info
TimeInfo = GetTime
End Property

Private Function GetTime() As Time_Info
'Track Time
Dim Gone As Long, Total As Long

Gone = SendUser(wmTrackTimeInfo, 0) / 1000
Total = SendUser(wmTrackTimeInfo, 1)

GetTime.TimeGone.Minutes = (Gone / 60)
GetTime.TimeGone.Seconds = (Gone Mod 60)

GetTime.TimeTotal.Minutes = (Total / 60)
GetTime.TimeTotal.Seconds = (Total Mod 60)

GetTime.TimeLeft.Minutes = (Total - Gone) \ 60
GetTime.TimeLeft.Seconds = (Total - Gone) Mod 60
On Error Resume Next
GetTime.Percent = CLng(((SendUser(wmTrackTimeInfo, 0)) / 1000) / (SendUser(wmTrackTimeInfo, 1)) * 100)

End Function

Private Function GetInfo() As Track_Info
Dim wFullName As String, WinTxt As String, Pos As Long
On Error Resume Next

'Get Full Playing Status
wFullName = Space$(255)
WinTxt = GetWindowText(hwndWinamp, wFullName, 256)

'Track No
Pos = InStr(1, wFullName, ".")
GetInfo.Number = Mid$(wFullName, 1, Pos)
wFullName = Mid$(wFullName, Pos + 2)

'Artist
Pos = InStr(1, wFullName, " - ")
GetInfo.Artist = Mid$(wFullName, 1, Pos - 1)
wFullName = Mid$(wFullName, Pos + 3)

' title
Pos = InStr(1, wFullName, " - Winamp")
GetInfo.Title = Mid$(wFullName, 1, Pos - 1)

'Current Status
Select Case SendUser(wmPlaybackStatus)
    Case 1: GetInfo.Status = "Playing"
    Case 3: GetInfo.Status = "Paused"
    Case Else: GetInfo.Status = "Stopped"
End Select

'Bitrate/Channels
GetInfo.Bitrate.KHZ = SendUser(wmTrackRateInfo, 0)
GetInfo.Bitrate.KBPS = SendUser(wmTrackRateInfo, 1)
GetInfo.Channels = SendUser(wmTrackRateInfo, 2)

'Playlist Count
GetInfo.ListCount = SendUser(wmPlaylistLength)

End Function


