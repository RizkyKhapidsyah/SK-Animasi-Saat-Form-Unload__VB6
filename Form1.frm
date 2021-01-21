VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Tutup"
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Created by Rizky Khapidsyah

Dim GotoVal As Long
Dim GoInto As Long

Sub StarTrek(frm As Form)

GotoVal = frm.Height / 2
For GoInto = 1 To GotoVal
    DoEvents
    frm.Height = frm.Height - 100
    frm.Top = (Screen.Height - frm.Height) \ 2
    If frm.Height <= 500 Then Exit For
Next GoInto

horiz:
frm.Height = 30
GotoVal = frm.Width / 2
    For GoInto = 1 To GotoVal
    DoEvents
    frm.Width = frm.Width - 100
    frm.Left = (Screen.Width - frm.Width) \ 2
    If frm.Width <= 2000 Then Exit For
Next GoInto
Unload Me
End Sub


