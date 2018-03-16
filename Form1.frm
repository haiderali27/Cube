VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Rubik's Cube"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   37
      Top             =   4260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   2
      Height          =   5640
      Left            =   4845
      ScaleHeight     =   372
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   372
      TabIndex        =   33
      Top             =   390
      Width           =   5640
      Begin VB.VScrollBar Rott2 
         Height          =   5145
         LargeChange     =   45
         Left            =   5295
         Max             =   180
         Min             =   -180
         SmallChange     =   15
         TabIndex        =   36
         Top             =   30
         Value           =   30
         Width           =   240
      End
      Begin VB.HScrollBar Rott1 
         Height          =   240
         LargeChange     =   45
         Left            =   15
         Max             =   180
         Min             =   -180
         SmallChange     =   15
         TabIndex        =   35
         Top             =   5295
         Value           =   -30
         Width           =   5115
      End
      Begin VB.CommandButton Command26 
         Enabled         =   0   'False
         Height          =   345
         Left            =   5160
         TabIndex        =   34
         Top             =   5205
         Width           =   360
      End
   End
   Begin VB.CommandButton Mini 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11460
      TabIndex        =   27
      Top             =   15
      Width           =   255
   End
   Begin VB.CommandButton Clos 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11730
      TabIndex        =   26
      Top             =   15
      Width           =   255
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   105
      Left            =   240
      TabIndex        =   24
      Top             =   7110
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Main Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   3030
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   4590
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   120
         Top             =   360
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   720
         Top             =   360
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3360
         TabIndex        =   23
         Text            =   "10"
         Top             =   2520
         Width           =   690
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Scramble Cube"
         Height          =   360
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   480
         TabIndex        =   28
         Top             =   720
         Width           =   4365
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   1200
         TabIndex        =   25
         Top             =   1680
         Width           =   4365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scramble Level :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1920
         TabIndex        =   22
         Top             =   2520
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Whole Cube Moves"
      ForeColor       =   &H00FFFFFF&
      Height          =   1560
      Left            =   60
      TabIndex        =   13
      Top             =   7380
      Width           =   4605
      Begin VB.CommandButton Command18 
         Caption         =   "Left"
         Height          =   360
         Left            =   2355
         TabIndex        =   19
         Top             =   1065
         Width           =   2160
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Down"
         Height          =   360
         Left            =   2355
         TabIndex        =   18
         Top             =   660
         Width           =   2160
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Back"
         Height          =   360
         Left            =   2355
         TabIndex        =   17
         Top             =   255
         Width           =   2160
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Right"
         Height          =   360
         Left            =   135
         TabIndex        =   16
         Top             =   1080
         Width           =   2160
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Top"
         Height          =   360
         Left            =   135
         TabIndex        =   15
         Top             =   660
         Width           =   2160
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Front"
         Height          =   360
         Left            =   135
         TabIndex        =   14
         Top             =   255
         Width           =   2160
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Face Moves"
      ForeColor       =   &H00FFFFFF&
      Height          =   2790
      Left            =   4755
      TabIndex        =   0
      Top             =   6150
      Width           =   5805
      Begin VB.CommandButton Command12 
         Caption         =   "Down (Clock-Wise)"
         Height          =   360
         Left            =   2355
         TabIndex        =   12
         Top             =   1920
         Width           =   2160
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Down (Anti Clock-Wise)"
         Height          =   360
         Left            =   2355
         TabIndex        =   11
         Top             =   2325
         Width           =   2160
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Back (Clock-Wise)"
         Height          =   360
         Left            =   2355
         TabIndex        =   10
         Top             =   1110
         Width           =   2160
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Back (Anti Clock-Wise)"
         Height          =   360
         Left            =   2355
         TabIndex        =   9
         Top             =   1515
         Width           =   2160
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Top (Clock-Wise)"
         Height          =   360
         Left            =   2355
         TabIndex        =   8
         Top             =   300
         Width           =   2160
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Top (Anti Clock-Wise)"
         Height          =   360
         Left            =   2355
         TabIndex        =   7
         Top             =   705
         Width           =   2160
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Left (Clock-Wise)"
         Height          =   360
         Left            =   135
         TabIndex        =   6
         Top             =   1920
         Width           =   2160
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Left (Anti Clock-Wise)"
         Height          =   360
         Left            =   135
         TabIndex        =   5
         Top             =   2325
         Width           =   2160
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Right (Clock-Wise)"
         Height          =   360
         Left            =   135
         TabIndex        =   4
         Top             =   1110
         Width           =   2160
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Right (Anti Clock-Wise)"
         Height          =   360
         Left            =   135
         TabIndex        =   3
         Top             =   1515
         Width           =   2160
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Front (Anti Clock-Wise)"
         Height          =   360
         Left            =   135
         TabIndex        =   2
         Top             =   705
         Width           =   2160
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Front (Clock-Wise)"
         Height          =   360
         Left            =   135
         TabIndex        =   1
         Top             =   300
         Width           =   2160
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   4125
      Left            =   105
      TabIndex        =   29
      Top             =   270
      Width           =   4605
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   3705
         TabIndex        =   32
         Top             =   2205
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Left            =   3720
         TabIndex        =   31
         Top             =   2220
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox Text2 
         Height          =   705
         Left            =   825
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   16
         Left            =   2565
         Picture         =   "Form1.frx":08CA
         Stretch         =   -1  'True
         Top             =   2775
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   10
         Left            =   2565
         Picture         =   "Form1.frx":0CD3
         Stretch         =   -1  'True
         Top             =   1740
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   11
         Left            =   2745
         Picture         =   "Form1.frx":10DC
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   12
         Left            =   2925
         Picture         =   "Form1.frx":14E5
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   13
         Left            =   2565
         Picture         =   "Form1.frx":18EE
         Stretch         =   -1  'True
         Top             =   2265
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   14
         Left            =   2745
         Picture         =   "Form1.frx":1CF7
         Stretch         =   -1  'True
         Top             =   2070
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   15
         Left            =   2925
         Picture         =   "Form1.frx":2100
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   17
         Left            =   2745
         Picture         =   "Form1.frx":2509
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   18
         Left            =   2925
         Picture         =   "Form1.frx":2912
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   480
         Index           =   1
         Left            =   1005
         Picture         =   "Form1.frx":2D1B
         Stretch         =   -1  'True
         Top             =   1875
         Width           =   480
      End
      Begin VB.Image imgPanel 
         Height          =   480
         Index           =   2
         Left            =   1530
         Picture         =   "Form1.frx":5D5D
         Stretch         =   -1  'True
         Top             =   1875
         Width           =   480
      End
      Begin VB.Image imgPanel 
         Height          =   480
         Index           =   3
         Left            =   2055
         Picture         =   "Form1.frx":8D9F
         Stretch         =   -1  'True
         Top             =   1875
         Width           =   480
      End
      Begin VB.Image imgPanel 
         Height          =   480
         Index           =   4
         Left            =   1005
         Picture         =   "Form1.frx":BDE1
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image imgPanel 
         Height          =   480
         Index           =   5
         Left            =   1530
         Picture         =   "Form1.frx":EE23
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image imgPanel 
         Height          =   480
         Index           =   6
         Left            =   2055
         Picture         =   "Form1.frx":11E65
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image imgPanel 
         Height          =   480
         Index           =   7
         Left            =   1005
         Picture         =   "Form1.frx":14EA7
         Stretch         =   -1  'True
         Top             =   2925
         Width           =   480
      End
      Begin VB.Image imgPanel 
         Height          =   480
         Index           =   8
         Left            =   1530
         Picture         =   "Form1.frx":17EE9
         Stretch         =   -1  'True
         Top             =   2925
         Width           =   480
      End
      Begin VB.Image imgPanel 
         Height          =   480
         Index           =   9
         Left            =   2055
         Picture         =   "Form1.frx":1AF2B
         Stretch         =   -1  'True
         Top             =   2925
         Width           =   480
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   37
         Left            =   1260
         Picture         =   "Form1.frx":1DF6D
         Stretch         =   -1  'True
         Top             =   1365
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   38
         Left            =   1815
         Picture         =   "Form1.frx":1E310
         Stretch         =   -1  'True
         Top             =   1365
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   39
         Left            =   2355
         Picture         =   "Form1.frx":1E6B3
         Stretch         =   -1  'True
         Top             =   1365
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   40
         Left            =   1110
         Picture         =   "Form1.frx":1EA56
         Stretch         =   -1  'True
         Top             =   1530
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   41
         Left            =   1650
         Picture         =   "Form1.frx":1EDF9
         Stretch         =   -1  'True
         Top             =   1530
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   42
         Left            =   2190
         Picture         =   "Form1.frx":1F19C
         Stretch         =   -1  'True
         Top             =   1530
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   43
         Left            =   945
         Picture         =   "Form1.frx":1F53F
         Stretch         =   -1  'True
         Top             =   1695
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   44
         Left            =   1485
         Picture         =   "Form1.frx":1F8E2
         Stretch         =   -1  'True
         Top             =   1695
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   45
         Left            =   2025
         Picture         =   "Form1.frx":1FC85
         Stretch         =   -1  'True
         Top             =   1695
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   46
         Left            =   1545
         Picture         =   "Form1.frx":20028
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   47
         Left            =   2085
         Picture         =   "Form1.frx":203CB
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   48
         Left            =   2625
         Picture         =   "Form1.frx":2076E
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   49
         Left            =   1395
         Picture         =   "Form1.frx":20B11
         Stretch         =   -1  'True
         Top             =   3675
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   50
         Left            =   1935
         Picture         =   "Form1.frx":20EB4
         Stretch         =   -1  'True
         Top             =   3675
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   51
         Left            =   2490
         Picture         =   "Form1.frx":21257
         Stretch         =   -1  'True
         Top             =   3675
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   52
         Left            =   1230
         Picture         =   "Form1.frx":215FA
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   53
         Left            =   1785
         Picture         =   "Form1.frx":2199D
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   135
         Index           =   54
         Left            =   2325
         Picture         =   "Form1.frx":21D40
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   675
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   36
         Left            =   210
         Picture         =   "Form1.frx":220E3
         Stretch         =   -1  'True
         Top             =   2715
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   30
         Left            =   210
         Picture         =   "Form1.frx":224EC
         Stretch         =   -1  'True
         Top             =   1695
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   29
         Left            =   390
         Picture         =   "Form1.frx":228F5
         Stretch         =   -1  'True
         Top             =   1515
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   28
         Left            =   570
         Picture         =   "Form1.frx":22CFE
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   33
         Left            =   210
         Picture         =   "Form1.frx":23107
         Stretch         =   -1  'True
         Top             =   2205
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   32
         Left            =   390
         Picture         =   "Form1.frx":23510
         Stretch         =   -1  'True
         Top             =   2025
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   31
         Left            =   570
         Picture         =   "Form1.frx":23919
         Stretch         =   -1  'True
         Top             =   1845
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   35
         Left            =   390
         Picture         =   "Form1.frx":23D22
         Stretch         =   -1  'True
         Top             =   2535
         Width           =   165
      End
      Begin VB.Image imgPanel 
         Height          =   675
         Index           =   34
         Left            =   570
         Picture         =   "Form1.frx":2412B
         Stretch         =   -1  'True
         Top             =   2355
         Width           =   165
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   2115
         Left            =   930
         Top             =   1335
         Width           =   2190
      End
      Begin VB.Image imgPanel 
         Height          =   420
         Index           =   21
         Left            =   2820
         Picture         =   "Form1.frx":24534
         Stretch         =   -1  'True
         Top             =   180
         Width           =   420
      End
      Begin VB.Image imgPanel 
         Height          =   420
         Index           =   20
         Left            =   3270
         Picture         =   "Form1.frx":27576
         Stretch         =   -1  'True
         Top             =   180
         Width           =   420
      End
      Begin VB.Image imgPanel 
         Height          =   420
         Index           =   19
         Left            =   3720
         Picture         =   "Form1.frx":2A5B8
         Stretch         =   -1  'True
         Top             =   180
         Width           =   420
      End
      Begin VB.Image imgPanel 
         Height          =   420
         Index           =   24
         Left            =   2820
         Picture         =   "Form1.frx":2D5FA
         Stretch         =   -1  'True
         Top             =   630
         Width           =   420
      End
      Begin VB.Image imgPanel 
         Height          =   420
         Index           =   23
         Left            =   3270
         Picture         =   "Form1.frx":3063C
         Stretch         =   -1  'True
         Top             =   630
         Width           =   420
      End
      Begin VB.Image imgPanel 
         Height          =   420
         Index           =   22
         Left            =   3720
         Picture         =   "Form1.frx":3367E
         Stretch         =   -1  'True
         Top             =   630
         Width           =   420
      End
      Begin VB.Image imgPanel 
         Height          =   420
         Index           =   27
         Left            =   2820
         Picture         =   "Form1.frx":366C0
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   420
      End
      Begin VB.Image imgPanel 
         Height          =   420
         Index           =   26
         Left            =   3270
         Picture         =   "Form1.frx":39702
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   420
      End
      Begin VB.Image imgPanel 
         Height          =   420
         Index           =   25
         Left            =   3720
         Picture         =   "Form1.frx":3C744
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   420
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim CubeA As String
Dim SolPos As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' Also used to make the form draggable
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&
Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ
Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte
Const DisplayErrorMsg = False
'3D Drawing

Private Type Frames
    X1 As Single
    Y1 As Single
    Z1 As Single
    X2 As Single
    Y2 As Single
    Z2 As Single
    X3 As Single
    Y3 As Single
    Z3 As Single
    X4 As Single
    Y4 As Single
    Z4 As Single
End Type

Private Type Pnt
    X As Single
    Y As Single
    Z As Single
End Type

Dim ori(54) As Frames
Dim dup(54) As Frames
Dim colo(400) As Long
Dim Pont(8) As Pnt


Private Sub Rotate(X, Y, q)
xd = X * Cos(q) - Y * Sin(q)
yd = X * Sin(q) + Y * Cos(q)
X = xd
Y = yd
End Sub

Private Sub DrawPolygon(Fram As Frames, Colour As Long)
On Error Resume Next
X1 = Fram.X1
Y1 = Fram.Y1
Z1 = Fram.Z1
X2 = Fram.X2
Y2 = Fram.Y2
Z2 = Fram.Z2
X3 = Fram.X3
Y3 = Fram.Y3
Z3 = Fram.Z3
X4 = Fram.X4
Y4 = Fram.Y4
Z4 = Fram.Z4
X1 = X1 * (1000 - Z1) / 1000
Y1 = Y1 * (1000 - Z1) / 1000
X2 = X2 * (1000 - Z2) / 1000
Y2 = Y2 * (1000 - Z2) / 1000
X3 = X3 * (1000 - Z3) / 1000
Y3 = Y3 * (1000 - Z3) / 1000
X4 = X4 * (1000 - Z4) / 1000
Y4 = Y4 * (1000 - Z4) / 1000




'If Abs(X2 - X1) > 0.1 Then
For xx = X1 To X2 Step Sgn(X2 - X1) * Abs(X2 - X1) / 30 + 0.0000001
Picture1.Line (xx, Y1 + (Y2 - Y1) * (xx - X1) / (X2 - X1))-(X4 + (X3 - X4) * (xx - X1) / (X2 - X1), Y4 + (Y3 - Y4) * (xx - X1) / (X2 - X1)), Colour
Next
'Else
'For yy = Y1 To Y2 Step Sgn(Y2 - Y1) * Abs(Y2 - Y1) / 400 + 0.0000001
'Line (Y1 + (yy - Y1) * (X2 - X1) / (Y2 - Y1), yy)-(X4 + (X3 - X4) * (yy - Y1) / (Y2 - Y1), Y4 + (Y3 - Y4) * (yy - Y1) / (Y2 - Y2)), Colour
'Next
Picture1.Line (X1, Y1)-(X2, Y2), 0
Picture1.Line (X2, Y2)-(X3, Y3), 0
Picture1.Line (X3, Y3)-(X4, Y4), 0
Picture1.Line (X4, Y4)-(X1, Y1), 0
'End If
End Sub

Private Sub DrawPolygon1(Fram As Frames, Colour As Long)
On Error Resume Next
X1 = Fram.X1
Y1 = Fram.Y1
Z1 = -Fram.Z1 + 300
X2 = Fram.X2
Y2 = Fram.Y2
Z2 = -Fram.Z2 + 300
X3 = Fram.X3
Y3 = Fram.Y3
Z3 = -Fram.Z3 + 300
X4 = Fram.X4
Y4 = Fram.Y4
Z4 = -Fram.Z4 + 300
X1 = X1 * (1000 - Z1) / 1000
Y1 = Y1 * (1000 - Z1) / 1000
X2 = X2 * (1000 - Z2) / 1000
Y2 = Y2 * (1000 - Z2) / 1000
X3 = X3 * (1000 - Z3) / 1000
Y3 = Y3 * (1000 - Z3) / 1000
X4 = X4 * (1000 - Z4) / 1000
Y4 = Y4 * (1000 - Z4) / 1000




'If Abs(X2 - X1) > 0.1 Then
For xx = X1 To X2 Step Sgn(X2 - X1) * Abs(X2 - X1) / 30 + 0.0000001
Picture1.Line (xx, Y1 + (Y2 - Y1) * (xx - X1) / (X2 - X1))-(X4 + (X3 - X4) * (xx - X1) / (X2 - X1), Y4 + (Y3 - Y4) * (xx - X1) / (X2 - X1)), Colour
Next
'Else
'For yy = Y1 To Y2 Step Sgn(Y2 - Y1) * Abs(Y2 - Y1) / 400 + 0.0000001
'Line (Y1 + (yy - Y1) * (X2 - X1) / (Y2 - Y1), yy)-(X4 + (X3 - X4) * (yy - Y1) / (Y2 - Y1), Y4 + (Y3 - Y4) * (yy - Y1) / (Y2 - Y2)), Colour
'Next
Picture1.Line (X1, Y1)-(X2, Y2), 0
Picture1.Line (X2, Y2)-(X3, Y3), 0
Picture1.Line (X3, Y3)-(X4, Y4), 0
Picture1.Line (X4, Y4)-(X1, Y1), 0
'End If
End Sub




Function SetDWORDValue(SubKey As String, Entry As String, Value As Long)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then
      rtn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4)
      If Not rtn = ERROR_SUCCESS Then
         If DisplayErrorMsg = True Then
            MsgBox ErrorMsg(rtn)
         End If
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user want errors displayed
         MsgBox ErrorMsg(rtn) 'display the error
      End If
   End If
End If

End Function
Function GetDWORDValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         GetDWORDValue = lBuffer  'return the value
      Else                        'otherwise, if the value couldnt be retreived
         GetDWORDValue = "Error"  'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants errors displayed
            MsgBox ErrorMsg(rtn)        'tell the user what was wrong
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetDWORDValue = "Error"        'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MsgBox ErrorMsg(rtn)        'tell the user what was wrong
      End If
   End If
End If

End Function

Function SetBinaryValue(SubKey As String, Entry As String, Value As String)
Dim i
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      lDataSize = Len(Value)
      ReDim ByteArray(lDataSize)
      For i = 1 To lDataSize
      ByteArray(i) = Asc(Mid$(Value, i, 1))
      Next
      rtn = RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if the was an error writting the value
         If DisplayErrorMsg = True Then 'if the user want errors displayed
            MsgBox ErrorMsg(rtn)        'display the error
         End If
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MsgBox ErrorMsg(rtn) 'display the error
      End If
   End If
End If

End Function

Function GetBinaryValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened
      lBufferSize = 1
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize) 'get the value from the registry
      sBuffer = Space(lBufferSize)
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         GetBinaryValue = sBuffer 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetBinaryValue = "Error" 'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants to errors displayed
            MsgBox ErrorMsg(rtn)  'display the error to the user
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetBinaryValue = "Error" 'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants to errors displayed
         MsgBox ErrorMsg(rtn)  'display the error to the user
      End If
   End If
End If

End Function
Function DeleteKey(KeyName As String)

Call ParseKey(KeyName, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, KeyName, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = RegDeleteKey(hKey, KeyName) 'delete the key
      rtn = RegCloseKey(hKey)  'close the key
   End If
End If

End Function

Function GetMainKeyHandle(MainKeyName As String) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
   
Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function

Function ErrorMsg(lErrorCode As Long) As String
    Dim GetErrorMsg
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1009, 1015
            GetErrorMsg = "The Registry Database is corrupt!"
       Case 2, 1010
            GetErrorMsg = "Bad Key Name"
       Case 1011
            GetErrorMsg = "Can't Open Key"
       Case 4, 1012
            GetErrorMsg = "Can't Read Key"
       Case 5
            GetErrorMsg = "Access to this key is denied"
       Case 1013
            GetErrorMsg = "Can't Write Key"
       Case 8, 14
            GetErrorMsg = "Out of memory"
       Case 87
            GetErrorMsg = "Invalid Parameter"
       Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case Else
            GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
End Select

End Function

Function GetStringValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      sBuffer = Space(255)     'make a buffer
      lBufferSize = Len(sBuffer)
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         sBuffer = Trim(sBuffer)
         GetStringValue = Left(sBuffer, Len(sBuffer) - 1) 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetStringValue = "Error" 'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants errors displayed then
            MsgBox ErrorMsg(rtn)  'tell the user what was wrong
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetStringValue = "Error"       'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants errors displayed then
         MsgBox ErrorMsg(rtn)        'tell the user what was wrong
      End If
   End If
End If

End Function

Private Sub ParseKey(KeyName As String, Keyhandle As Long)
    
rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname

If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + KeyName 'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(KeyName)
   KeyName = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1)) 'seperate the Keyname
   KeyName = Right(KeyName, Len(KeyName) - rtn)
End If

End Sub
Function CreateKey(SubKey As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegCreateKey(MainKeyHandle, SubKey, hKey) 'create the key
   If rtn = ERROR_SUCCESS Then 'if the key was created then
      rtn = RegCloseKey(hKey)  'close the key
   End If
End If

End Function
Function SetStringValue(SubKey As String, Entry As String, Value As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value)) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         If DisplayErrorMsg = True Then 'if the user wants errors displayed
            MsgBox ErrorMsg(rtn)        'display the error
         End If
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MsgBox ErrorMsg(rtn)        'display the error
      End If
   End If
End If

End Function










Private Function OpenIt(ToOpen As String)
    ShellExecute &O0, "Open", ToOpen, &O0, &O0, 1
End Function

Private Sub AlwaysOnTop(EnabledOrDisabled, FormID As Object)
     If EnabledOrDisabled = "Enabled" Then SetWindowPos FormID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
     If EnabledOrDisabled = "Disabled" Then SetWindowPos FormID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub DoDrag(TheForm As Form)
' TheForm:  The form you want to start dragging
    
    If TheForm.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage TheForm.hwnd, &HA1, 2, 0&
    End If
End Sub



Function ConvertCube(Cube As String) As String
Dim F, R, B, L, U, D As String
Dim s As String
F = Mid(CubeA, 5, 1)
R = Mid(CubeA, 14, 1)
B = Mid(CubeA, 23, 1)
L = Mid(CubeA, 32, 1)
U = Mid(CubeA, 41, 1)
D = Mid(CubeA, 50, 1)
s = ""
s = s + FindPiece(CubeA, U + F) + " "
s = s + FindPiece(CubeA, U + R) + " "
s = s + FindPiece(CubeA, U + B) + " "
s = s + FindPiece(CubeA, U + L) + " "

s = s + FindPiece(CubeA, D + F) + " "
s = s + FindPiece(CubeA, D + R) + " "
s = s + FindPiece(CubeA, D + B) + " "
s = s + FindPiece(CubeA, D + L) + " "

s = s + FindPiece(CubeA, F + R) + " "
s = s + FindPiece(CubeA, F + L) + " "
s = s + FindPiece(CubeA, B + R) + " "
s = s + FindPiece(CubeA, B + L) + " "

s = s + FindPiece(CubeA, U + F + R) + " "
s = s + FindPiece(CubeA, U + R + B) + " "
s = s + FindPiece(CubeA, U + B + L) + " "
s = s + FindPiece(CubeA, U + L + F) + " "

s = s + FindPiece(CubeA, D + R + F) + " "
s = s + FindPiece(CubeA, D + F + L) + " "
s = s + FindPiece(CubeA, D + L + B) + " "
s = s + FindPiece(CubeA, D + B + R)

ConvertCube = s
End Function

Sub RotateFace(Cube As String, Face As String)
Select Case Face
Case "f" ' Rotate front face Clock-wise
temp$ = Mid(Cube, 2, 1)
Mid(Cube, 2, 1) = Mid(Cube, 4, 1)
Mid(Cube, 4, 1) = Mid(Cube, 8, 1)
Mid(Cube, 8, 1) = Mid(Cube, 6, 1)
Mid(Cube, 6, 1) = temp$
temp$ = Mid(Cube, 1, 1)
Mid(Cube, 1, 1) = Mid(Cube, 7, 1)
Mid(Cube, 7, 1) = Mid(Cube, 9, 1)
Mid(Cube, 9, 1) = Mid(Cube, 3, 1)
Mid(Cube, 3, 1) = temp$

temp$ = Mid(Cube, 43, 1)
Mid(Cube, 43, 1) = Mid(Cube, 36, 1)
Mid(Cube, 36, 1) = Mid(Cube, 54, 1)
Mid(Cube, 54, 1) = Mid(Cube, 10, 1)
Mid(Cube, 10, 1) = temp$
temp$ = Mid(Cube, 44, 1)
Mid(Cube, 44, 1) = Mid(Cube, 33, 1)
Mid(Cube, 33, 1) = Mid(Cube, 53, 1)
Mid(Cube, 53, 1) = Mid(Cube, 13, 1)
Mid(Cube, 13, 1) = temp$
temp$ = Mid(Cube, 45, 1)
Mid(Cube, 45, 1) = Mid(Cube, 30, 1)
Mid(Cube, 30, 1) = Mid(Cube, 52, 1)
Mid(Cube, 52, 1) = Mid(Cube, 16, 1)
Mid(Cube, 16, 1) = temp$

Case "r" ' Rotate front face Clock-wise
temp$ = Mid(Cube, 11, 1)
Mid(Cube, 11, 1) = Mid(Cube, 13, 1)
Mid(Cube, 13, 1) = Mid(Cube, 17, 1)
Mid(Cube, 17, 1) = Mid(Cube, 15, 1)
Mid(Cube, 15, 1) = temp$
temp$ = Mid(Cube, 10, 1)
Mid(Cube, 10, 1) = Mid(Cube, 16, 1)
Mid(Cube, 16, 1) = Mid(Cube, 18, 1)
Mid(Cube, 18, 1) = Mid(Cube, 12, 1)
Mid(Cube, 12, 1) = temp$

temp$ = Mid(Cube, 45, 1)
Mid(Cube, 45, 1) = Mid(Cube, 9, 1)
Mid(Cube, 9, 1) = Mid(Cube, 48, 1)
Mid(Cube, 48, 1) = Mid(Cube, 19, 1)
Mid(Cube, 19, 1) = temp$
temp$ = Mid(Cube, 42, 1)
Mid(Cube, 42, 1) = Mid(Cube, 6, 1)
Mid(Cube, 6, 1) = Mid(Cube, 51, 1)
Mid(Cube, 51, 1) = Mid(Cube, 22, 1)
Mid(Cube, 22, 1) = temp$
temp$ = Mid(Cube, 39, 1)
Mid(Cube, 39, 1) = Mid(Cube, 3, 1)
Mid(Cube, 3, 1) = Mid(Cube, 54, 1)
Mid(Cube, 54, 1) = Mid(Cube, 25, 1)
Mid(Cube, 25, 1) = temp$

Case "l" ' Rotate front face Clock-wise
temp$ = Mid(Cube, 30, 1)
Mid(Cube, 30, 1) = Mid(Cube, 28, 1)
Mid(Cube, 28, 1) = Mid(Cube, 34, 1)
Mid(Cube, 34, 1) = Mid(Cube, 36, 1)
Mid(Cube, 36, 1) = temp$
temp$ = Mid(Cube, 29, 1)
Mid(Cube, 29, 1) = Mid(Cube, 31, 1)
Mid(Cube, 31, 1) = Mid(Cube, 35, 1)
Mid(Cube, 35, 1) = Mid(Cube, 33, 1)
Mid(Cube, 33, 1) = temp$

temp$ = Mid(Cube, 1, 1)
Mid(Cube, 1, 1) = Mid(Cube, 37, 1)
Mid(Cube, 37, 1) = Mid(Cube, 27, 1)
Mid(Cube, 27, 1) = Mid(Cube, 52, 1)
Mid(Cube, 52, 1) = temp$
temp$ = Mid(Cube, 4, 1)
Mid(Cube, 4, 1) = Mid(Cube, 40, 1)
Mid(Cube, 40, 1) = Mid(Cube, 24, 1)
Mid(Cube, 24, 1) = Mid(Cube, 49, 1)
Mid(Cube, 49, 1) = temp$
temp$ = Mid(Cube, 7, 1)
Mid(Cube, 7, 1) = Mid(Cube, 43, 1)
Mid(Cube, 43, 1) = Mid(Cube, 21, 1)
Mid(Cube, 21, 1) = Mid(Cube, 46, 1)
Mid(Cube, 46, 1) = temp$

Case "t" ' Rotate front face Clock-wise
temp$ = Mid(Cube, 37, 1)
Mid(Cube, 37, 1) = Mid(Cube, 43, 1)
Mid(Cube, 43, 1) = Mid(Cube, 45, 1)
Mid(Cube, 45, 1) = Mid(Cube, 39, 1)
Mid(Cube, 39, 1) = temp$
temp$ = Mid(Cube, 40, 1)
Mid(Cube, 40, 1) = Mid(Cube, 44, 1)
Mid(Cube, 44, 1) = Mid(Cube, 42, 1)
Mid(Cube, 42, 1) = Mid(Cube, 38, 1)
Mid(Cube, 38, 1) = temp$

temp$ = Mid(Cube, 1, 1)
Mid(Cube, 1, 1) = Mid(Cube, 10, 1)
Mid(Cube, 10, 1) = Mid(Cube, 19, 1)
Mid(Cube, 19, 1) = Mid(Cube, 28, 1)
Mid(Cube, 28, 1) = temp$
temp$ = Mid(Cube, 29, 1)
Mid(Cube, 29, 1) = Mid(Cube, 2, 1)
Mid(Cube, 2, 1) = Mid(Cube, 11, 1)
Mid(Cube, 11, 1) = Mid(Cube, 20, 1)
Mid(Cube, 20, 1) = temp$
temp$ = Mid(Cube, 3, 1)
Mid(Cube, 3, 1) = Mid(Cube, 12, 1)
Mid(Cube, 12, 1) = Mid(Cube, 21, 1)
Mid(Cube, 21, 1) = Mid(Cube, 30, 1)
Mid(Cube, 30, 1) = temp$

Case "b" ' Rotate front face Clock-wise
temp$ = Mid(Cube, 21, 1)
Mid(Cube, 21, 1) = Mid(Cube, 19, 1)
Mid(Cube, 19, 1) = Mid(Cube, 25, 1)
Mid(Cube, 25, 1) = Mid(Cube, 27, 1)
Mid(Cube, 27, 1) = temp$
temp$ = Mid(Cube, 20, 1)
Mid(Cube, 20, 1) = Mid(Cube, 22, 1)
Mid(Cube, 22, 1) = Mid(Cube, 26, 1)
Mid(Cube, 26, 1) = Mid(Cube, 24, 1)
Mid(Cube, 24, 1) = temp$

temp$ = Mid(Cube, 37, 1)
Mid(Cube, 37, 1) = Mid(Cube, 12, 1)
Mid(Cube, 12, 1) = Mid(Cube, 48, 1)
Mid(Cube, 48, 1) = Mid(Cube, 34, 1)
Mid(Cube, 34, 1) = temp$
temp$ = Mid(Cube, 38, 1)
Mid(Cube, 38, 1) = Mid(Cube, 15, 1)
Mid(Cube, 15, 1) = Mid(Cube, 47, 1)
Mid(Cube, 47, 1) = Mid(Cube, 31, 1)
Mid(Cube, 31, 1) = temp$
temp$ = Mid(Cube, 39, 1)
Mid(Cube, 39, 1) = Mid(Cube, 18, 1)
Mid(Cube, 18, 1) = Mid(Cube, 46, 1)
Mid(Cube, 46, 1) = Mid(Cube, 28, 1)
Mid(Cube, 28, 1) = temp$

Case "d" ' Rotate front face Clock-wise
temp$ = Mid(Cube, 46, 1)
Mid(Cube, 46, 1) = Mid(Cube, 48, 1)
Mid(Cube, 48, 1) = Mid(Cube, 54, 1)
Mid(Cube, 54, 1) = Mid(Cube, 52, 1)
Mid(Cube, 52, 1) = temp$
temp$ = Mid(Cube, 47, 1)
Mid(Cube, 47, 1) = Mid(Cube, 51, 1)
Mid(Cube, 51, 1) = Mid(Cube, 53, 1)
Mid(Cube, 53, 1) = Mid(Cube, 49, 1)
Mid(Cube, 49, 1) = temp$

temp$ = Mid(Cube, 16, 1)
Mid(Cube, 16, 1) = Mid(Cube, 7, 1)
Mid(Cube, 7, 1) = Mid(Cube, 34, 1)
Mid(Cube, 34, 1) = Mid(Cube, 25, 1)
Mid(Cube, 25, 1) = temp$
temp$ = Mid(Cube, 8, 1)
Mid(Cube, 8, 1) = Mid(Cube, 35, 1)
Mid(Cube, 35, 1) = Mid(Cube, 26, 1)
Mid(Cube, 26, 1) = Mid(Cube, 17, 1)
Mid(Cube, 17, 1) = temp$
temp$ = Mid(Cube, 9, 1)
Mid(Cube, 9, 1) = Mid(Cube, 36, 1)
Mid(Cube, 36, 1) = Mid(Cube, 27, 1)
Mid(Cube, 27, 1) = Mid(Cube, 18, 1)
Mid(Cube, 18, 1) = temp$
End Select
End Sub

Private Sub Clos_Click()
Unload Me
Unload Form2
End Sub

Private Sub Command1_Click()
Dim Fra(21) As Integer
Fra(1) = 1
Fra(2) = 2
Fra(3) = 3
Fra(4) = 4
Fra(5) = 5
Fra(6) = 6
Fra(7) = 7
Fra(8) = 8
Fra(9) = 9
Fra(10) = 10
Fra(11) = 13
Fra(12) = 16
Fra(13) = 43
Fra(14) = 44
Fra(15) = 45
Fra(16) = 30
Fra(17) = 33
Fra(18) = 36
Fra(19) = 52
Fra(20) = 53
Fra(21) = 54
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).X1, ori(Fra(j)).Y1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).X2, ori(Fra(j)).Y2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).X3, ori(Fra(j)).Y3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).X4, ori(Fra(j)).Y4, 3.14159265258979 * 22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Front Face Clock-wise"
RotateFace CubeA, "f"
ViewCube CubeA
Rott1_Scroll
End Sub

Private Sub Command10_Click()
Dim Fra(21) As Integer
Fra(1) = 19
Fra(2) = 20
Fra(3) = 21
Fra(4) = 22
Fra(5) = 23
Fra(6) = 24
Fra(7) = 25
Fra(8) = 26
Fra(9) = 27
Fra(10) = 12
Fra(11) = 15
Fra(12) = 18
Fra(13) = 37
Fra(14) = 38
Fra(15) = 39
Fra(16) = 28
Fra(17) = 31
Fra(18) = 34
Fra(19) = 46
Fra(20) = 47
Fra(21) = 48
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).X1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).X2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).X3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).X4, 3.14159265258979 * 22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Back Face Clock-wise"
List1.AddItem "10"
RotateFace CubeA, "b"
ViewCube CubeA
End Sub

Private Sub Command11_Click()
Dim Fra(21) As Integer
Fra(1) = 46
Fra(2) = 47
Fra(3) = 48
Fra(4) = 49
Fra(5) = 50
Fra(6) = 51
Fra(7) = 52
Fra(8) = 53
Fra(9) = 54
Fra(10) = 7
Fra(11) = 8
Fra(12) = 9
Fra(13) = 16
Fra(14) = 17
Fra(15) = 18
Fra(16) = 25
Fra(17) = 26
Fra(18) = 27
Fra(19) = 34
Fra(20) = 35
Fra(21) = 36
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Z1, ori(Fra(j)).X1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z2, ori(Fra(j)).X2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z3, ori(Fra(j)).X3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z4, ori(Fra(j)).X4, 3.14159265258979 * 22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Bottom Face Anti Clock-wise"
List1.AddItem "11"
RotateFace CubeA, "d"
RotateFace CubeA, "d"
RotateFace CubeA, "d"
ViewCube CubeA
End Sub

Private Sub Command12_Click()
Dim Fra(21) As Integer
Fra(1) = 46
Fra(2) = 47
Fra(3) = 48
Fra(4) = 49
Fra(5) = 50
Fra(6) = 51
Fra(7) = 52
Fra(8) = 53
Fra(9) = 54
Fra(10) = 7
Fra(11) = 8
Fra(12) = 9
Fra(13) = 16
Fra(14) = 17
Fra(15) = 18
Fra(16) = 25
Fra(17) = 26
Fra(18) = 27
Fra(19) = 34
Fra(20) = 35
Fra(21) = 36
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Z1, ori(Fra(j)).X1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z2, ori(Fra(j)).X2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z3, ori(Fra(j)).X3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z4, ori(Fra(j)).X4, 3.14159265258979 * -22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Bottom Face Clock-wise"
List1.AddItem "12"
RotateFace CubeA, "d"
ViewCube CubeA
End Sub

Private Sub Command13_Click()
RotateFace CubeA, "f"
RotateFace CubeA, "b"
RotateFace CubeA, "b"
RotateFace CubeA, "b"

temp$ = Mid(CubeA, 17, 1)
Mid(CubeA, 17, 1) = Mid(CubeA, 42, 1)
Mid(CubeA, 42, 1) = Mid(CubeA, 29, 1)
Mid(CubeA, 29, 1) = Mid(CubeA, 49, 1)
Mid(CubeA, 49, 1) = temp$
temp$ = Mid(CubeA, 14, 1)
Mid(CubeA, 14, 1) = Mid(CubeA, 41, 1)
Mid(CubeA, 41, 1) = Mid(CubeA, 32, 1)
Mid(CubeA, 32, 1) = Mid(CubeA, 50, 1)
Mid(CubeA, 50, 1) = temp$
temp$ = Mid(CubeA, 11, 1)
Mid(CubeA, 11, 1) = Mid(CubeA, 40, 1)
Mid(CubeA, 40, 1) = Mid(CubeA, 35, 1)
Mid(CubeA, 35, 1) = Mid(CubeA, 51, 1)
Mid(CubeA, 51, 1) = temp$

ViewCube CubeA
End Sub

Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = "Click to turn whole cube clock-wise along Front face"
End Sub

Private Sub Command14_Click()
RotateFace CubeA, "t"
RotateFace CubeA, "d"
RotateFace CubeA, "d"
RotateFace CubeA, "d"

temp$ = Mid(CubeA, 4, 1)
Mid(CubeA, 4, 1) = Mid(CubeA, 13, 1)
Mid(CubeA, 13, 1) = Mid(CubeA, 22, 1)
Mid(CubeA, 22, 1) = Mid(CubeA, 31, 1)
Mid(CubeA, 31, 1) = temp$
temp$ = Mid(CubeA, 5, 1)
Mid(CubeA, 5, 1) = Mid(CubeA, 14, 1)
Mid(CubeA, 14, 1) = Mid(CubeA, 23, 1)
Mid(CubeA, 23, 1) = Mid(CubeA, 32, 1)
Mid(CubeA, 32, 1) = temp$
temp$ = Mid(CubeA, 6, 1)
Mid(CubeA, 6, 1) = Mid(CubeA, 15, 1)
Mid(CubeA, 15, 1) = Mid(CubeA, 24, 1)
Mid(CubeA, 24, 1) = Mid(CubeA, 33, 1)
Mid(CubeA, 33, 1) = temp$

ViewCube CubeA
End Sub

Private Sub Command14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = "Click to turn whole cube clock-wise along Top face"
End Sub

Private Sub Command15_Click()
RotateFace CubeA, "r"
RotateFace CubeA, "l"
RotateFace CubeA, "l"
RotateFace CubeA, "l"

temp$ = Mid(CubeA, 44, 1)
Mid(CubeA, 44, 1) = Mid(CubeA, 8, 1)
Mid(CubeA, 8, 1) = Mid(CubeA, 47, 1)
Mid(CubeA, 47, 1) = Mid(CubeA, 20, 1)
Mid(CubeA, 20, 1) = temp$
temp$ = Mid(CubeA, 41, 1)
Mid(CubeA, 41, 1) = Mid(CubeA, 5, 1)
Mid(CubeA, 5, 1) = Mid(CubeA, 50, 1)
Mid(CubeA, 50, 1) = Mid(CubeA, 23, 1)
Mid(CubeA, 23, 1) = temp$
temp$ = Mid(CubeA, 38, 1)
Mid(CubeA, 38, 1) = Mid(CubeA, 2, 1)
Mid(CubeA, 2, 1) = Mid(CubeA, 53, 1)
Mid(CubeA, 53, 1) = Mid(CubeA, 26, 1)
Mid(CubeA, 26, 1) = temp$

ViewCube CubeA
End Sub

Private Sub Command15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = "Click to turn whole cube clock-wise along Right face"
End Sub

Private Sub Command16_Click()
Command13_Click
Command13_Click
Command13_Click
End Sub

Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = "Click to turn whole cube clock-wise along Back face"
End Sub

Private Sub Command17_Click()
Command14_Click
Command14_Click
Command14_Click
End Sub

Private Sub Command17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = "Click to turn whole cube clock-wise along Bottom face"
End Sub

Private Sub Command18_Click()
Command15_Click
Command15_Click
Command15_Click
End Sub

Private Sub Command18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = "Click to turn whole cube clock-wise along Left face"
End Sub

Private Sub Command19_Click()
For i = 1 To Val(Text1.Text)
Randomize
n = Round(Rnd * 2342123) Mod 12
Select Case (n + 1)
Case 1
RotateFace CubeA, "f"
Case 2
RotateFace CubeA, "f"
RotateFace CubeA, "f"
RotateFace CubeA, "f"
Case 3
RotateFace CubeA, "r"
Case 4
RotateFace CubeA, "r"
RotateFace CubeA, "r"
RotateFace CubeA, "r"
Case 5
RotateFace CubeA, "l"
Case 6
RotateFace CubeA, "l"
RotateFace CubeA, "l"
RotateFace CubeA, "l"
Case 7
RotateFace CubeA, "b"
Case 8
RotateFace CubeA, "b"
RotateFace CubeA, "b"
RotateFace CubeA, "b"
Case 9
RotateFace CubeA, "t"
Case 10
RotateFace CubeA, "t"
RotateFace CubeA, "t"
RotateFace CubeA, "t"
Case 11
RotateFace CubeA, "d"
Case 12
RotateFace CubeA, "d"
RotateFace CubeA, "d"
RotateFace CubeA, "d"
End Select
Next
DoEvents
ViewCube CubeA
Label2.Caption = "Cube is now Scrambled."
End Sub

Private Sub Command2_Click()
Dim Fra(21) As Integer
Fra(1) = 1
Fra(2) = 2
Fra(3) = 3
Fra(4) = 4
Fra(5) = 5
Fra(6) = 6
Fra(7) = 7
Fra(8) = 8
Fra(9) = 9
Fra(10) = 10
Fra(11) = 13
Fra(12) = 16
Fra(13) = 43
Fra(14) = 44
Fra(15) = 45
Fra(16) = 30
Fra(17) = 33
Fra(18) = 36
Fra(19) = 52
Fra(20) = 53
Fra(21) = 54
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).X1, ori(Fra(j)).Y1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).X2, ori(Fra(j)).Y2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).X3, ori(Fra(j)).Y3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).X4, ori(Fra(j)).Y4, 3.14159265258979 * -22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Front Face Anti Clock-wise"
List1.AddItem "2"
RotateFace CubeA, "f"
RotateFace CubeA, "f"
RotateFace CubeA, "f"
ViewCube CubeA
Rott1_Scroll
End Sub








Private Sub Command24_Click()

End Sub




Private Sub Command3_Click()
Dim Fra(21) As Integer
Fra(1) = 10
Fra(2) = 11
Fra(3) = 12
Fra(4) = 13
Fra(5) = 14
Fra(6) = 15
Fra(7) = 16
Fra(8) = 17
Fra(9) = 18
Fra(10) = 3
Fra(11) = 6
Fra(12) = 9
Fra(13) = 19
Fra(14) = 22
Fra(15) = 25
Fra(16) = 39
Fra(17) = 42
Fra(18) = 45
Fra(19) = 48
Fra(20) = 51
Fra(21) = 54
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).Z1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).Z2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).Z3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).Z4, 3.14159265258979 * 22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Right Face Anti Clock-wise"
List1.AddItem "3"
RotateFace CubeA, "r"
RotateFace CubeA, "r"
RotateFace CubeA, "r"
ViewCube CubeA
End Sub

Private Sub Command4_Click()
Dim Fra(21) As Integer
Fra(1) = 10
Fra(2) = 11
Fra(3) = 12
Fra(4) = 13
Fra(5) = 14
Fra(6) = 15
Fra(7) = 16
Fra(8) = 17
Fra(9) = 18
Fra(10) = 3
Fra(11) = 6
Fra(12) = 9
Fra(13) = 19
Fra(14) = 22
Fra(15) = 25
Fra(16) = 39
Fra(17) = 42
Fra(18) = 45
Fra(19) = 48
Fra(20) = 51
Fra(21) = 54
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).Z1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).Z2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).Z3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).Z4, 3.14159265258979 * -22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Right Face Clock-wise"
List1.AddItem "4"
RotateFace CubeA, "r"
ViewCube CubeA
End Sub

Private Sub Command5_Click()
Dim Fra(21) As Integer
Fra(1) = 28
Fra(2) = 29
Fra(3) = 30
Fra(4) = 31
Fra(5) = 32
Fra(6) = 33
Fra(7) = 34
Fra(8) = 35
Fra(9) = 36
Fra(10) = 1
Fra(11) = 4
Fra(12) = 7
Fra(13) = 37
Fra(14) = 40
Fra(15) = 43
Fra(16) = 46
Fra(17) = 49
Fra(18) = 52
Fra(19) = 21
Fra(20) = 24
Fra(21) = 27
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).Z1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).Z2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).Z3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).Z4, 3.14159265258979 * -22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Left Face Anti Clock-wise"
List1.AddItem "5"
RotateFace CubeA, "l"
RotateFace CubeA, "l"
RotateFace CubeA, "l"
ViewCube CubeA
End Sub

Private Sub Command6_Click()
Dim Fra(21) As Integer
Fra(1) = 28
Fra(2) = 29
Fra(3) = 30
Fra(4) = 31
Fra(5) = 32
Fra(6) = 33
Fra(7) = 34
Fra(8) = 35
Fra(9) = 36
Fra(10) = 1
Fra(11) = 4
Fra(12) = 7
Fra(13) = 37
Fra(14) = 40
Fra(15) = 43
Fra(16) = 46
Fra(17) = 49
Fra(18) = 52
Fra(19) = 21
Fra(20) = 24
Fra(21) = 27
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).Z1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).Z2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).Z3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).Z4, 3.14159265258979 * 22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Left Face Clock-wise"
List1.AddItem "6"
RotateFace CubeA, "l"
ViewCube CubeA
End Sub

Private Sub Command7_Click()
Dim Fra(21) As Integer
Fra(1) = 37
Fra(2) = 38
Fra(3) = 39
Fra(4) = 40
Fra(5) = 41
Fra(6) = 42
Fra(7) = 43
Fra(8) = 44
Fra(9) = 45
Fra(10) = 1
Fra(11) = 2
Fra(12) = 3
Fra(13) = 10
Fra(14) = 11
Fra(15) = 12
Fra(16) = 19
Fra(17) = 20
Fra(18) = 21
Fra(19) = 28
Fra(20) = 29
Fra(21) = 30
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Z1, ori(Fra(j)).X1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z2, ori(Fra(j)).X2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z3, ori(Fra(j)).X3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z4, ori(Fra(j)).X4, 3.14159265258979 * -22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Top Face Anti Clock-wise"
List1.AddItem "7"
RotateFace CubeA, "t"
RotateFace CubeA, "t"
RotateFace CubeA, "t"
ViewCube CubeA
End Sub

Private Sub Command8_Click()
Dim Fra(21) As Integer
Fra(1) = 37
Fra(2) = 38
Fra(3) = 39
Fra(4) = 40
Fra(5) = 41
Fra(6) = 42
Fra(7) = 43
Fra(8) = 44
Fra(9) = 45
Fra(10) = 1
Fra(11) = 2
Fra(12) = 3
Fra(13) = 10
Fra(14) = 11
Fra(15) = 12
Fra(16) = 19
Fra(17) = 20
Fra(18) = 21
Fra(19) = 28
Fra(20) = 29
Fra(21) = 30
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Z1, ori(Fra(j)).X1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z2, ori(Fra(j)).X2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z3, ori(Fra(j)).X3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z4, ori(Fra(j)).X4, 3.14159265258979 * 22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Top Face Clock-wise"
List1.AddItem "8"
RotateFace CubeA, "t"
ViewCube CubeA
End Sub

Private Sub Command9_Click()
Dim Fra(21) As Integer
Fra(1) = 19
Fra(2) = 20
Fra(3) = 21
Fra(4) = 22
Fra(5) = 23
Fra(6) = 24
Fra(7) = 25
Fra(8) = 26
Fra(9) = 27
Fra(10) = 12
Fra(11) = 15
Fra(12) = 18
Fra(13) = 37
Fra(14) = 38
Fra(15) = 39
Fra(16) = 28
Fra(17) = 31
Fra(18) = 34
Fra(19) = 46
Fra(20) = 47
Fra(21) = 48
For ii = 1 To 4
Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).X1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).X2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).X3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).X4, 3.14159265258979 * -22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    DrawCube
'    Picture1.Scale (-300, -300)-(700, 700)
'    DrawCube1
Next

Label2.Caption = "Rotate Back Face Anti Clock-wise"
List1.AddItem "9"
RotateFace CubeA, "b"
RotateFace CubeA, "b"
RotateFace CubeA, "b"
ViewCube CubeA
End Sub

Private Sub Cuber_DepthDone(ByVal Depth As Integer)
ProgressBar1.Value = Depth * 100 / 13
End Sub

Private Sub Cuber_SolutionFound(ByVal SolutionString As String, ByVal NumQtrTurns As Integer)
On Error Resume Next
Text2.Text = SolutionString
List2.Clear
For i = 1 To Len(SolutionString) Step 3
List2.AddItem Mid(SolutionString, i, 2)
Next
End Sub

Private Sub Cuber_TablesInitProgress(ByVal PercentDone As Integer)
ProgressBar1.Value = PercentDone
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim za As Integer
If KeyCode = vbKeyF4 And Shift = vbAltMask Then
Form_Unload za
End
End If
End Sub

Private Sub Form_Load()
'MousePointer = 11
CubeA = "RRRRRRRRRYYYYYYYYYPPPPPPPPPWWWWWWWWWBBBBBBBBBGGGGGGGGG"
'ViewCube CubeA
'Cuber.InitTables
'MousePointer = 0
s$ = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\RubikCube", "Cube")
If s$ = "Error" Then
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\RubikCube"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\RubikCube", "Cube", CubeA
Else
CubeA = s$
End If
Picture1.Scale (-700, -700)-(300, 300)
Rott1_Change
End Sub

Sub ViewCube(Cube As String)
Cube = UCase(Cube)

'Front
For i = 1 To 9
Select Case Mid(Cube, i, 1)
Case "G"
imgPanel(i).Picture = Form2.imgPanel(0).Picture
Case "Y"
imgPanel(i).Picture = Form2.imgPanel(1).Picture
Case "B"
imgPanel(i).Picture = Form2.imgPanel(2).Picture
Case "W"
imgPanel(i).Picture = Form2.imgPanel(3).Picture
Case "R"
imgPanel(i).Picture = Form2.imgPanel(4).Picture
Case "P"
imgPanel(i).Picture = Form2.imgPanel(5).Picture
End Select
Next

'right
For i = 10 To 18
Select Case Mid(Cube, i, 1)
Case "G"
imgPanel(i).Picture = Form2.imgPanel(6).Picture
Case "Y"
imgPanel(i).Picture = Form2.imgPanel(7).Picture
Case "B"
imgPanel(i).Picture = Form2.imgPanel(8).Picture
Case "W"
imgPanel(i).Picture = Form2.imgPanel(9).Picture
Case "R"
imgPanel(i).Picture = Form2.imgPanel(10).Picture
Case "P"
imgPanel(i).Picture = Form2.imgPanel(11).Picture
End Select
Next

'Back
For i = 19 To 27
Select Case Mid(Cube, i, 1)
Case "G"
imgPanel(i).Picture = Form2.imgPanel(0).Picture
Case "Y"
imgPanel(i).Picture = Form2.imgPanel(1).Picture
Case "B"
imgPanel(i).Picture = Form2.imgPanel(2).Picture
Case "W"
imgPanel(i).Picture = Form2.imgPanel(3).Picture
Case "R"
imgPanel(i).Picture = Form2.imgPanel(4).Picture
Case "P"
imgPanel(i).Picture = Form2.imgPanel(5).Picture
End Select
Next

'Left
For i = 28 To 36
Select Case Mid(Cube, i, 1)
Case "G"
imgPanel(i).Picture = Form2.imgPanel(6).Picture
Case "Y"
imgPanel(i).Picture = Form2.imgPanel(7).Picture
Case "B"
imgPanel(i).Picture = Form2.imgPanel(8).Picture
Case "W"
imgPanel(i).Picture = Form2.imgPanel(9).Picture
Case "R"
imgPanel(i).Picture = Form2.imgPanel(10).Picture
Case "P"
imgPanel(i).Picture = Form2.imgPanel(11).Picture
End Select
Next

'Top and Bottom
For i = 37 To 54
Select Case Mid(Cube, i, 1)
Case "G"
imgPanel(i).Picture = Form2.imgPanel(12).Picture
Case "Y"
imgPanel(i).Picture = Form2.imgPanel(13).Picture
Case "B"
imgPanel(i).Picture = Form2.imgPanel(14).Picture
Case "W"
imgPanel(i).Picture = Form2.imgPanel(15).Picture
Case "R"
imgPanel(i).Picture = Form2.imgPanel(16).Picture
Case "P"
imgPanel(i).Picture = Form2.imgPanel(17).Picture
End Select
Next

Rott1_Scroll
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\RubikCube", "Cube", CubeA
End Sub

Private Sub imgPanel_Click(Index As Integer)
If Command23.Caption = "Exit Edit Cube" Then
If Mid(CubeA, Index, 1) = "B" Then Mid(CubeA, Index, 1) = "P": GoTo 1000
If Mid(CubeA, Index, 1) = "W" Then Mid(CubeA, Index, 1) = "B"
If Mid(CubeA, Index, 1) = "Y" Then Mid(CubeA, Index, 1) = "W"
If Mid(CubeA, Index, 1) = "R" Then Mid(CubeA, Index, 1) = "Y"
If Mid(CubeA, Index, 1) = "G" Then Mid(CubeA, Index, 1) = "R"
If Mid(CubeA, Index, 1) = "P" Then Mid(CubeA, Index, 1) = "G"
Else
Beep
End If
1000 ViewCube CubeA
End Sub

Private Sub Rott1_Change()
Init3D
For i = 1 To 54
Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
Next
For i = 1 To 54
Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
Next
DoEvents
Picture1.Cls
Picture1.Scale (-700, -700)-(300, 300)
DrawCube
Picture1.Scale (-300, -300)-(700, 700)
DrawCube1
End Sub

Private Sub Rott1_Scroll()
Init3D
For i = 1 To 54
Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
Next
For i = 1 To 54
Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
Next
DoEvents
Picture1.Cls
Picture1.Scale (-700, -700)-(300, 300)
DrawCube
Picture1.Scale (-300, -300)-(700, 700)
DrawCube1
End Sub

Private Sub Rott2_Change()
Init3D
For i = 1 To 54
Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
Next
For i = 1 To 54
Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
Next
DoEvents
Picture1.Cls
Picture1.Scale (-700, -700)-(300, 300)
DrawCube
Picture1.Scale (-300, -300)-(700, 700)
DrawCube1
End Sub

Private Sub Rott2_Scroll()
Init3D
For i = 1 To 54
Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
Next
For i = 1 To 54
Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
Next
DoEvents
Picture1.Cls
Picture1.Scale (-700, -700)-(300, 300)
DrawCube
Picture1.Scale (-300, -300)-(700, 700)
DrawCube1
End Sub



Private Sub Mini_Click()
WindowState = 1
End Sub

Private Sub ProgressBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = "Shows Thinking Progress."
End Sub


Private Sub Timer1_Timer()
SolPos = SolPos - 1
n = Val(List1.List(SolPos))
Select Case (n)
Case 1
Command2_Click
Case 2
Command1_Click
Case 3
Command4_Click
Case 4
Command3_Click
Case 5
Command6_Click
Case 6
Command5_Click
Case 7
Command8_Click
Case 8
Command7_Click
Case 9
Command10_Click
Case 10
Command9_Click
Case 11
Command12_Click
Case 12
Command11_Click
End Select
If SolPos = 0 Then
    List1.Clear
    Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
SolPos = SolPos + 1
n = Val(List2.List(SolPos))
Select Case (n)
Case 1
Command2_Click
Case 2
Command1_Click
Case 3
Command4_Click
Case 4
Command3_Click
Case 5
Command6_Click
Case 6
Command5_Click
Case 7
Command8_Click
Case 8
Command7_Click
Case 9
Command10_Click
Case 10
Command9_Click
Case 11
Command12_Click
Case 12
Command11_Click
End Select
If SolPos = 0 Then
    List2.Clear
    Timer2.Enabled = False
End If
End Sub

Function FindPiece(Cube As String, Piece As String) As String
Dim Pos As String
Select Case Len(Piece)
Case 1 'Middle Piece
If Mid(Cube, 5, 1) = Piece Then Pos = "F"
If Mid(Cube, 14, 1) = Piece Then Pos = "R"
If Mid(Cube, 5, 1) = Piece Then Pos = "F"
If Mid(Cube, 5, 1) = Piece Then Pos = "F"
If Mid(Cube, 5, 1) = Piece Then Pos = "F"
If Mid(Cube, 5, 1) = Piece Then Pos = "F"
Case 2 'Edge Piece
m = Mid(Piece, 1, 1)
n = Mid(Piece, 2, 1)

If Mid(Cube, 44, 1) = m And Mid(Cube, 2, 1) = n Then Pos = "UF"
If Mid(Cube, 44, 1) = n And Mid(Cube, 2, 1) = m Then Pos = "FU"

If Mid(Cube, 13, 1) = m And Mid(Cube, 6, 1) = n Then Pos = "RF"
If Mid(Cube, 13, 1) = n And Mid(Cube, 6, 1) = m Then Pos = "FR"

If Mid(Cube, 53, 1) = m And Mid(Cube, 8, 1) = n Then Pos = "DF"
If Mid(Cube, 53, 1) = n And Mid(Cube, 8, 1) = m Then Pos = "FD"

If Mid(Cube, 33, 1) = m And Mid(Cube, 4, 1) = n Then Pos = "LF"
If Mid(Cube, 33, 1) = n And Mid(Cube, 4, 1) = m Then Pos = "FL"

If Mid(Cube, 42, 1) = m And Mid(Cube, 11, 1) = n Then Pos = "UR"
If Mid(Cube, 42, 1) = n And Mid(Cube, 11, 1) = m Then Pos = "RU"

If Mid(Cube, 40, 1) = m And Mid(Cube, 29, 1) = n Then Pos = "UL"
If Mid(Cube, 40, 1) = n And Mid(Cube, 29, 1) = m Then Pos = "LU"

If Mid(Cube, 51, 1) = m And Mid(Cube, 17, 1) = n Then Pos = "DR"
If Mid(Cube, 51, 1) = n And Mid(Cube, 17, 1) = m Then Pos = "RD"

If Mid(Cube, 49, 1) = m And Mid(Cube, 35, 1) = n Then Pos = "DL"
If Mid(Cube, 49, 1) = n And Mid(Cube, 35, 1) = m Then Pos = "LD"

If Mid(Cube, 20, 1) = m And Mid(Cube, 38, 1) = n Then Pos = "BU"
If Mid(Cube, 20, 1) = n And Mid(Cube, 38, 1) = m Then Pos = "UB"

If Mid(Cube, 22, 1) = m And Mid(Cube, 15, 1) = n Then Pos = "BR"
If Mid(Cube, 22, 1) = n And Mid(Cube, 15, 1) = m Then Pos = "RB"

If Mid(Cube, 26, 1) = m And Mid(Cube, 47, 1) = n Then Pos = "BD"
If Mid(Cube, 26, 1) = n And Mid(Cube, 47, 1) = m Then Pos = "DB"

If Mid(Cube, 24, 1) = m And Mid(Cube, 31, 1) = n Then Pos = "BL"
If Mid(Cube, 24, 1) = n And Mid(Cube, 31, 1) = m Then Pos = "LB"

Case 3 'Corner Piece
Dim a(5) As String
Dim B(5) As String
Dim c(5) As String
a(0) = Mid(Piece, 1, 1): B(0) = Mid(Piece, 2, 1): c(0) = Mid(Piece, 3, 1)
a(1) = a(0): B(1) = c(0): c(1) = B(0)
a(2) = B(0): B(2) = a(0): c(2) = c(0)
a(3) = B(0): B(3) = c(0): c(3) = a(0)
a(4) = c(0): B(4) = a(0): c(4) = B(0)
a(5) = c(0): B(5) = B(0): c(5) = a(0)
For i = 0 To 5
If Mid(Cube, 45, 1) = a(i) And Mid(Cube, 3, 1) = B(i) And Mid(Cube, 10, 1) = c(i) Then Pos = "UFR"
If Mid(Cube, 43, 1) = a(i) And Mid(Cube, 1, 1) = B(i) And Mid(Cube, 30, 1) = c(i) Then Pos = "UFL"
If Mid(Cube, 39, 1) = a(i) And Mid(Cube, 19, 1) = B(i) And Mid(Cube, 12, 1) = c(i) Then Pos = "UBR"
If Mid(Cube, 37, 1) = a(i) And Mid(Cube, 21, 1) = B(i) And Mid(Cube, 28, 1) = c(i) Then Pos = "UBL"
If Mid(Cube, 54, 1) = a(i) And Mid(Cube, 9, 1) = B(i) And Mid(Cube, 16, 1) = c(i) Then Pos = "DFR"
If Mid(Cube, 52, 1) = a(i) And Mid(Cube, 7, 1) = B(i) And Mid(Cube, 36, 1) = c(i) Then Pos = "DFL"
If Mid(Cube, 48, 1) = a(i) And Mid(Cube, 25, 1) = B(i) And Mid(Cube, 18, 1) = c(i) Then Pos = "DBR"
If Mid(Cube, 46, 1) = a(i) And Mid(Cube, 27, 1) = B(i) And Mid(Cube, 34, 1) = c(i) Then Pos = "DBL"
If Pos <> "" Then Exit For
Next
If i = 1 Then Pos = Mid(Pos, 1, 1) + Mid(Pos, 3, 1) + Mid(Pos, 2, 1)
If i = 2 Then Pos = Mid(Pos, 2, 1) + Mid(Pos, 1, 1) + Mid(Pos, 3, 1)
If i = 3 Then Pos = Mid(Pos, 3, 1) + Mid(Pos, 1, 1) + Mid(Pos, 2, 1)
If i = 4 Then Pos = Mid(Pos, 2, 1) + Mid(Pos, 3, 1) + Mid(Pos, 1, 1)
If i = 5 Then Pos = Mid(Pos, 3, 1) + Mid(Pos, 2, 1) + Mid(Pos, 1, 1)
End Select
FindPiece = Pos
End Function

Private Sub Init3D()

For i = 1 To 9
ori(i).Z1 = -150
ori(i).Z2 = -150
ori(i).Z3 = -150
ori(i).Z4 = -150
ori(i).X1 = -150 + ((i - 1) Mod 3) * 100
ori(i).X2 = -50 + ((i - 1) Mod 3) * 100
ori(i).X3 = -50 + ((i - 1) Mod 3) * 100
ori(i).X4 = -150 + ((i - 1) Mod 3) * 100
ori(i).Y1 = -150 + ((i - 1) \ 3) * 100
ori(i).Y2 = -150 + ((i - 1) \ 3) * 100
ori(i).Y3 = -50 + ((i - 1) \ 3) * 100
ori(i).Y4 = -50 + ((i - 1) \ 3) * 100
Next

For i = 10 To 18
ori(i).X1 = 150
ori(i).X2 = 150
ori(i).X3 = 150
ori(i).X4 = 150
ori(i).Z1 = -150 + ((i - 1) Mod 3) * 100
ori(i).Z2 = -50 + ((i - 1) Mod 3) * 100
ori(i).Z3 = -50 + ((i - 1) Mod 3) * 100
ori(i).Z4 = -150 + ((i - 1) Mod 3) * 100
ori(i).Y1 = -150 + ((i - 10) \ 3) * 100
ori(i).Y2 = -150 + ((i - 10) \ 3) * 100
ori(i).Y3 = -50 + ((i - 10) \ 3) * 100
ori(i).Y4 = -50 + ((i - 10) \ 3) * 100
Next

For i = 19 To 27
ori(i).Z1 = 150
ori(i).Z2 = 150
ori(i).Z3 = 150
ori(i).Z4 = 150
ori(i).X1 = 150 - ((i - 1) Mod 3) * 100
ori(i).X2 = 50 - ((i - 1) Mod 3) * 100
ori(i).X3 = 50 - ((i - 1) Mod 3) * 100
ori(i).X4 = 150 - ((i - 1) Mod 3) * 100
ori(i).Y1 = -150 + ((i - 19) \ 3) * 100
ori(i).Y2 = -150 + ((i - 19) \ 3) * 100
ori(i).Y3 = -50 + ((i - 19) \ 3) * 100
ori(i).Y4 = -50 + ((i - 19) \ 3) * 100
Next


For i = 28 To 36
ori(i).X1 = -150
ori(i).X2 = -150
ori(i).X3 = -150
ori(i).X4 = -150
ori(i).Z1 = 150 - ((i - 1) Mod 3) * 100
ori(i).Z2 = 50 - ((i - 1) Mod 3) * 100
ori(i).Z3 = 50 - ((i - 1) Mod 3) * 100
ori(i).Z4 = 150 - ((i - 1) Mod 3) * 100
ori(i).Y1 = -150 + ((i - 28) \ 3) * 100
ori(i).Y2 = -150 + ((i - 28) \ 3) * 100
ori(i).Y3 = -50 + ((i - 28) \ 3) * 100
ori(i).Y4 = -50 + ((i - 28) \ 3) * 100
Next

For i = 37 To 45
ori(i).Y1 = -150
ori(i).Y2 = -150
ori(i).Y3 = -150
ori(i).Y4 = -150
ori(i).X1 = -150 + ((i - 1) Mod 3) * 100
ori(i).X2 = -50 + ((i - 1) Mod 3) * 100
ori(i).X3 = -50 + ((i - 1) Mod 3) * 100
ori(i).X4 = -150 + ((i - 1) Mod 3) * 100
ori(i).Z1 = 150 - ((i - 37) \ 3) * 100
ori(i).Z2 = 150 - ((i - 37) \ 3) * 100
ori(i).Z3 = 50 - ((i - 37) \ 3) * 100
ori(i).Z4 = 50 - ((i - 37) \ 3) * 100
Next

For i = 46 To 54
ori(i).Y1 = 150
ori(i).Y2 = 150
ori(i).Y3 = 150
ori(i).Y4 = 150
ori(i).X1 = -150 + ((i - 1) Mod 3) * 100
ori(i).X2 = -50 + ((i - 1) Mod 3) * 100
ori(i).X3 = -50 + ((i - 1) Mod 3) * 100
ori(i).X4 = -150 + ((i - 1) Mod 3) * 100
ori(i).Z1 = 150 - ((i - 46) \ 3) * 100
ori(i).Z2 = 150 - ((i - 46) \ 3) * 100
ori(i).Z3 = 50 - ((i - 46) \ 3) * 100
ori(i).Z4 = 50 - ((i - 46) \ 3) * 100
Next

For i = 1 To Len(CubeA)
Select Case Mid(CubeA, i, 1)
Case "R"
colo(i) = RGB(255, 0, 0)
Case "Y"
colo(i) = RGB(255, 255, 0)
Case "P"
colo(i) = &H80FF&
Case "W"
colo(i) = RGB(255, 255, 255)
Case "B"
colo(i) = RGB(50, 50, 200)
Case "G"
colo(i) = RGB(50, 200, 50)
End Select
Next
End Sub

Private Sub DrawCube()
For i = 1 To 54
List3.AddItem Str(FrDepth(ori(i))) & " " & Str(i)
Next
For i = 1 To 54
j = Val(Right(List3.List(i), 3))
DrawPolygon ori(j), colo(j)
Next
DoEvents
List3.Clear
End Sub

Private Sub DrawCube1()
Dim ord(6) As Integer
tmp = -1000
For i = 5 To 54 Step 9
If tmp < FrDepth(ori(i)) Then
ord(1) = i
tmp = FrDepth(ori(i))
End If
Next

tmp = -1000
For i = 5 To 54 Step 9
If tmp < FrDepth(ori(i)) And i <> ord(1) Then
ord(2) = i
tmp = FrDepth(ori(i))
End If
Next

tmp = -1000
For i = 5 To 54 Step 9
If tmp < FrDepth(ori(i)) And i <> ord(1) And i <> ord(2) Then
ord(3) = i
tmp = FrDepth(ori(i))
End If
Next

tmp = -1000
For i = 5 To 54 Step 9
If tmp < FrDepth(ori(i)) And i <> ord(1) And i <> ord(2) And i <> ord(3) Then
ord(4) = i
tmp = FrDepth(ori(i))
End If
Next

tmp = -1000
For i = 5 To 54 Step 9
If tmp < FrDepth(ori(i)) And i <> ord(1) And i <> ord(2) And i <> ord(3) And i <> ord(4) Then
ord(5) = i
tmp = FrDepth(ori(i))
End If
Next

tmp = -1000
For i = 5 To 54 Step 9
If tmp < FrDepth(ori(i)) And i <> ord(1) And i <> ord(2) And i <> ord(3) And i <> ord(4) And i <> ord(5) Then
ord(6) = i
tmp = FrDepth(ori(i))
End If
Next

For i = 4 To 6
xx = (ord(i) + 4) / 9
For j = ((xx - 1) * 9) + 1 To ((xx - 1) * 9) + 9
DrawPolygon1 ori(j), colo((j))
Next j
Next
End Sub

Private Function FrDepth(Fram As Frames) As Double
xx = (Fram.X1 + Fram.X2 + Fram.X3 + Fram.X4) / 4
yy = (Fram.Y1 + Fram.Y2 + Fram.Y3 + Fram.Y4) / 4
zz = (Fram.Z1 + Fram.Z2 + Fram.Z3 + Fram.Z4) / 4
FrDepth = xx ^ 2 + yy ^ 2 + (zz - 600) ^ 2
End Function
