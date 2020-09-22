VERSION 5.00
Begin VB.Form frmDbl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Double-Checking Information"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "frmDbl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "Please Wait..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmDbl.frx":1CCA
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmDbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a, b, c, d, e, f, g

Private Sub cmdOK_Click()
'hide form
Me.Visible = False
End Sub

Private Sub Form_Load()
'load registry settings
loadit
End Sub

Sub loadit()
'get registry settings
a = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Time Zones\Ekaterinburg", "DL")
b = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Time Zones\Ekaterinburg", "TZ")
c = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Time Zones\Ekaterinburg", "ID")
'get frmmain information
d = frmMain.Name2.Caption
e = frmMain.ID2.Caption
f = frmMain.Serial2.Caption
'make sure they are the same info
If a <> d Then g = g + 1
If b <> e Then g = g + 1
If c <> f Then g = g + 1
'if they aren't the same then error
If g <> 0 Then Label1.Caption = "An error has occurred.  Please stop causing errors or exit the program immediately.  Thank you."
cmdOK.Caption = "&OK"
cmdOK.Enabled = True
'if they are the same then give the ok text
If g = 0 Or Empty Then
Label1.Caption = "WOW!  someone actually registered..."
End If
End Sub
