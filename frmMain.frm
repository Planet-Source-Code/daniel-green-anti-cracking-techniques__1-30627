VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Software Protection"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3450
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCheckSerial 
      Caption         =   "Check Information"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Serial 
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox ID 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox Username 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "Dan Green"
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   $"frmMain.frx":1CCA
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Serial2 
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label ID2 
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Name2 
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Serial"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "User ID"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I wrote this to try to defend my software against
'any and all crackers.  Use this code freely if you
'want, just give me some credit.
'Basically what is does is generate a unique user id,
'then take 3 codes from the user-provided name along with
'user id, then combine it, check it with 2 registry points,
'recheck it.  that's it, just a little way to kill crackers.
'The end-product serial number turns out to be very, very long,
'which makes it harder to crack; and if they patch the program
'to accept any serial, they must patch every registry point as well!
'================
'Dan Green=2001=www.morphedmedia.com=liquidmotion@juno.com
'======:)========
Option Explicit
Dim a, b, c, d, e, f, g, h, i, j, k As String, l As String, m, n, o, p

Private Sub cmdCheckSerial_Click()
'makes sure all variables & such are empty
clearit
'generate 3 encryption keys (thank you W.G. Griffiths!)
a = crypt.KeyGen(Username, ID, 1)
b = crypt.KeyGen(Username, ID, 2)
c = crypt.KeyGen(Username, ID, 3)
'set variables to zero
e = 0
f = 0
h = 0
'start a loop to generate our serial from previous 3 encryption keys
Do Until h > Len(i)
e = e + 1
d = Left(a, e)
f = f + 1
g = Left(b, f)
h = h + 1
i = Left(c, h)
j = j & d & g & i
m = j
Loop
'general saving for other forms usage
Name2.Caption = Username
ID2 = ID
Serial2 = Serial
'1st registry save to deter against hacks
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Time Zones\Ekaterinburg", "DL", Username
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Time Zones\Ekaterinburg", "TZ", ID
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Time Zones\Ekaterinburg", "ID", Serial
'call 2nd registry save to deter against hacks
saveit
'if the serial is wrong then show the error box
If Serial <> m Then
frmErr.Show
Else
'if the serial is right then show the ok box
If Serial = m Then
frmDbl.Show
End If
End If
End Sub

Private Sub Form_Load()
'this generates the User ID #, which is unique to every machine
k = GetSettingString(HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion", "Plus! VersionNumber")
l = GetSettingString(HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
ID = crypt.KeyGen(l, k, 1)
clearit
End Sub

Sub saveit()
'the 2nd registry save to deter hacks
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Subscription Folder", "a", Username
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Subscription Folder", "b", ID
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Subscription Folder", "c", Serial
End Sub

Private Sub Form_Unload(Cancel As Integer)
'makes sure all variables & such are empty
clearit
'unload the forms and end program
Unload frmDbl
Unload frmErr
Unload frmMain
End
End Sub

Sub clearit()
'makes sure all variables & such are empty
a = Empty
b = Empty
c = Empty
d = Empty
e = Empty
f = Empty
g = Empty
h = Empty
o = Empty
j = Empty
k = Empty
l = Empty
m = Empty
n = Empty
o = Empty
p = Empty
Name2 = ""
ID2 = ""
Serial2 = ""
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Time Zones\Ekaterinburg", "DL", ""
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Time Zones\Ekaterinburg", "TZ", ""
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Time Zones\Ekaterinburg", "ID", ""
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Subscription Folder", "www.morphedmedia.com", ""
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Subscription Folder", "is", ""
SaveSettingString HKEY_LOCAL_MACHINE, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Subscription Folder", "the best", ""
End Sub
