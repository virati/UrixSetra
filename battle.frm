VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      BOFAction       =   1  'BOF
      Caption         =   "Guard"
      Connect         =   "FoxPro 3.0;"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\Vfp98"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2280
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   "table1"
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enemy: Guard"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command9 
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Height          =   435
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   135
      End
      Begin VB.CommandButton Command7 
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Height          =   975
         Left            =   1200
         TabIndex        =   5
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   975
         Left            =   840
         TabIndex        =   4
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Height          =   615
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "MG PWR:"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Class:"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub
