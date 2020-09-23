VERSION 5.00
Begin VB.Form fAbout 
   Caption         =   "About"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1605
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&Ok"
      Height          =   405
      Left            =   870
      TabIndex        =   0
      Top             =   1155
      Width           =   1635
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "lord_illogical@yahoo.com"
      Height          =   195
      Index           =   2
      Left            =   705
      TabIndex        =   3
      Top             =   645
      Width           =   1800
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "By Eric Sanford aka Lord_illogical"
      Height          =   195
      Index           =   1
      Left            =   405
      TabIndex        =   2
      Top             =   390
      Width           =   2370
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "Falling Snow example"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   135
      Width           =   1530
   End
End
Attribute VB_Name = "fAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCount  As Integer


Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    mCount = mCount + 1
    btnOK.Caption = "Close in " & (5 - mCount)
    If mCount > 5 Then
        Unload Me
    End If
End Sub
