VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4680
      Top             =   4920
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   -120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   4560
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3960
      Width           =   615
   End
   Begin MCI.MMControl mmc 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   3960
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   873
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "E:\Avseq06.dat"
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
mmc.From = mmc.Length
mmc.Command = "stop"


cmd.Filter = "*.*|*.*"
cmd.ShowOpen
mmc.FileName = cmd.FileName
mmc.hWndDisplay = Frame1.hWnd
mmc.Command = "open"
mmc.Command = "play"
HScroll1.Max = mmc.Length
Label1.Caption = mmc.Length
mmc.Shareable = True
Timer1.Enabled = True

End Sub

Private Sub HScroll1_Change()
mmc.From = HScroll1.Value
mmc.Command = "play"
End Sub

Private Sub Timer1_Timer()
If HScroll1.Value = Label1.Caption Then
HScroll1.Value = 0
Else
HScroll1.Value = mmc.Position
End If
End Sub
