VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2655
   ScaleWidth      =   3600
   Begin VB.CommandButton Command1 
      Caption         =   """Give"""
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Would you like a cup of coffee??"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   0
      Picture         =   "FORM1.frx":0000
      Top             =   0
      Width           =   3600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Sub Command1_Click()
    mciSendString "Set CDAudio Door Open Wait", _
        0&, 0&, 0&
        MsgBox "Now you have place to put it. Go make it now"
End Sub


