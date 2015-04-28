VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   240
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   3720
      Width           =   135
   End
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   4440
      Picture         =   "picsfrm.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   960
      Picture         =   "picsfrm.frx":20802
      ScaleHeight     =   3195
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
