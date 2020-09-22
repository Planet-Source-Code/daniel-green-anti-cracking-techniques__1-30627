VERSION 5.00
Begin VB.Form frmErr 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Error!"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "frmErr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "An error has occurred.  Please stop causing errors or exit the program immediately.  Thank you."
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form is used instead of a message box because
'a "cracker" can easily take out a msgbox command

Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub
