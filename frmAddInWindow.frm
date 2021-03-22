VERSION 5.00
Begin VB.Form frmAddInWindow 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "IDE actions logs"
   ClientHeight    =   2784
   ClientLeft      =   2172
   ClientTop       =   1932
   ClientWidth     =   8952
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2784
   ScaleWidth      =   8952
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtActions 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1908
      Left            =   144
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   5724
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   2208
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddInWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Sub cmdClose_Click()
    Connect.HideWindow
End Sub

Private Sub Form_Load()
    Me.Move GetSetting(App.Title, "Settings", "WindowLeft", Screen.Width - Me.Width - 200), GetSetting(App.Title, "Settings", "WindowTop", 0), GetSetting(App.Title, "Settings", "WindowWidth", Me.Width), GetSetting(App.Title, "Settings", "WindowHeight", Me.Height)
End Sub

Private Sub Form_Resize()
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 200, Me.ScaleHeight - cmdClose.Height - 200
    txtActions.Move 0, 0, Me.ScaleWidth, cmdClose.Top - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "WindowLeft", Me.Left
    SaveSetting App.Title, "Settings", "WindowTop", Me.Top
    SaveSetting App.Title, "Settings", "WindowWidth", Me.Width
    SaveSetting App.Title, "Settings", "WindowHeight", Me.Height
End Sub
