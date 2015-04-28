VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11760
   LinkTopic       =   "Form2"
   ScaleHeight     =   7755
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   9255
      Left            =   0
      Picture         =   "battle2.frx":0000
      ScaleHeight     =   9195
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton Command5 
         Caption         =   "X"
         Height          =   255
         Left            =   11400
         TabIndex        =   80
         Top             =   0
         Width           =   255
      End
      Begin VB.Timer Timer9 
         Interval        =   1000
         Left            =   3360
         Top             =   2400
      End
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   3360
         Top             =   2760
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   120
         Top             =   480
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   11280
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   78
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   0
         TabIndex        =   75
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         Caption         =   "Ability"
         Height          =   2055
         Left            =   3480
         TabIndex        =   67
         Top             =   5520
         Visible         =   0   'False
         Width           =   3135
         Begin VB.ComboBox Combo8 
            Height          =   315
            Left            =   840
            TabIndex        =   74
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   840
            TabIndex        =   73
            Top             =   1320
            Width           =   1455
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   840
            TabIndex        =   72
            Top             =   960
            Width           =   1455
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            ItemData        =   "battle2.frx":140E2A
            Left            =   720
            List            =   "battle2.frx":140E3A
            TabIndex        =   70
            Text            =   "------------"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label20 
            BackColor       =   &H00000000&
            Caption         =   "Ability 2"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label19 
            BackColor       =   &H00000000&
            Caption         =   "Ability"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Equip"
         Height          =   2055
         Left            =   3480
         TabIndex        =   57
         Top             =   5520
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton Command2 
            Caption         =   "Equip"
            Height          =   255
            Left            =   1080
            TabIndex        =   68
            Top             =   1680
            Width           =   855
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "battle2.frx":140E66
            Left            =   840
            List            =   "battle2.frx":140E70
            TabIndex        =   65
            Top             =   960
            Width           =   1575
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            ItemData        =   "battle2.frx":140E83
            Left            =   840
            List            =   "battle2.frx":140E8D
            TabIndex        =   64
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "battle2.frx":140EA6
            Left            =   840
            List            =   "battle2.frx":140EB0
            TabIndex        =   63
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "battle2.frx":140ED1
            Left            =   840
            List            =   "battle2.frx":140EDE
            TabIndex        =   62
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label25 
            BackColor       =   &H00000000&
            Caption         =   "Weapon:"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label23 
            BackColor       =   &H00000000&
            Caption         =   "Helmet:"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label22 
            BackColor       =   &H00000000&
            Caption         =   "Shield:"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label21 
            BackColor       =   &H00000000&
            Caption         =   "Ammo:"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1320
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "Inventory"
         Height          =   2055
         Left            =   3480
         TabIndex        =   55
         Top             =   5520
         Visible         =   0   'False
         Width           =   3135
         Begin VB.ListBox List1 
            Height          =   1620
            ItemData        =   "battle2.frx":140EF7
            Left            =   120
            List            =   "battle2.frx":140F04
            TabIndex        =   56
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Equip"
         Height          =   2055
         Left            =   3480
         TabIndex        =   46
         Top             =   5520
         Visible         =   0   'False
         Width           =   3135
         Begin VB.Label Label17 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   720
            TabIndex        =   54
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label16 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   720
            TabIndex        =   53
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label15 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   720
            TabIndex        =   52
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label14 
            BackColor       =   &H00000000&
            Caption         =   "Ammo:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label13 
            BackColor       =   &H00000000&
            Caption         =   "Shield:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label12 
            BackColor       =   &H00000000&
            Caption         =   "Helmet:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label11 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   48
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label10 
            BackColor       =   &H00000000&
            Caption         =   "Weapon:"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture12 
         Height          =   135
         Left            =   3120
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   45
         Top             =   7560
         Width           =   135
      End
      Begin VB.PictureBox Picture11 
         Height          =   135
         Left            =   3120
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   44
         Top             =   7440
         Width           =   135
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H00C0C0C0&
         Height          =   135
         Left            =   3120
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   43
         Top             =   7320
         Width           =   135
      End
      Begin VB.PictureBox Picture22 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   255
         TabIndex        =   39
         Top             =   5160
         Width           =   255
      End
      Begin VB.PictureBox Picture23 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1920
         ScaleHeight     =   495
         ScaleWidth      =   735
         TabIndex        =   38
         Top             =   5520
         Width           =   735
      End
      Begin VB.PictureBox Picture24 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   37
         Top             =   5520
         Width           =   375
      End
      Begin VB.PictureBox Picture25 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   36
         Top             =   5520
         Width           =   375
      End
      Begin VB.PictureBox Picture26 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   1920
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   35
         Top             =   6240
         Width           =   255
      End
      Begin VB.PictureBox Picture27 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2040
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   34
         Top             =   6000
         Width           =   495
      End
      Begin VB.PictureBox Picture28 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2400
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   33
         Top             =   6240
         Width           =   255
      End
      Begin VB.PictureBox Picture33 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   2880
         ScaleHeight     =   855
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   6000
         Width           =   255
      End
      Begin VB.Timer Timer8 
         Interval        =   850
         Left            =   600
         Top             =   4200
      End
      Begin VB.Timer Timer7 
         Left            =   4200
         Top             =   4920
      End
      Begin VB.Timer Timer6 
         Interval        =   1000
         Left            =   1080
         Top             =   4200
      End
      Begin VB.Timer Timer5 
         Left            =   4680
         Top             =   4920
      End
      Begin VB.Timer Timer4 
         Interval        =   1000
         Left            =   11040
         Top             =   480
      End
      Begin VB.Timer Timer1 
         Interval        =   850
         Left            =   120
         Top             =   4200
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Change"
         Height          =   1935
         Left            =   4080
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.PictureBox Picture18 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   720
            ScaleHeight     =   375
            ScaleWidth      =   255
            TabIndex        =   31
            Top             =   120
            Width           =   255
         End
         Begin VB.PictureBox Picture19 
            BackColor       =   &H0000FF00&
            Height          =   495
            Left            =   480
            ScaleHeight     =   435
            ScaleWidth      =   675
            TabIndex        =   30
            Top             =   480
            Width           =   735
         End
         Begin VB.PictureBox Picture20 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1200
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   29
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox Picture21 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   28
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox Picture29 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   480
            ScaleHeight     =   615
            ScaleWidth      =   255
            TabIndex        =   27
            Top             =   1200
            Width           =   255
         End
         Begin VB.PictureBox Picture30 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   600
            ScaleHeight     =   255
            ScaleWidth      =   495
            TabIndex        =   26
            Top             =   960
            Width           =   495
         End
         Begin VB.PictureBox Picture31 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   960
            ScaleHeight     =   615
            ScaleWidth      =   255
            TabIndex        =   25
            Top             =   1200
            Width           =   255
         End
         Begin VB.PictureBox Picture34 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   1440
            ScaleHeight     =   975
            ScaleWidth      =   255
            TabIndex        =   24
            Top             =   840
            Width           =   255
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar4 
         Height          =   2655
         Left            =   4080
         TabIndex        =   20
         Top             =   1920
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   4683
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         Picture         =   "battle2.frx":140F1F
         ScaleHeight     =   2175
         ScaleWidth      =   1095
         TabIndex        =   19
         Top             =   5040
         Width           =   1095
      End
      Begin VB.PictureBox Picture7 
         Height          =   255
         Left            =   1920
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   14
         Top             =   6960
         Width           =   255
      End
      Begin VB.PictureBox Picture6 
         Height          =   255
         Left            =   2280
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   13
         Top             =   6960
         Width           =   255
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   2520
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   12
         Top             =   6960
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         Height          =   255
         Left            =   1680
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   11
         Top             =   6960
         Width           =   255
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Command"
         Height          =   1815
         Left            =   6720
         TabIndex        =   4
         Top             =   5760
         Width           =   4815
         Begin VB.CommandButton Command4 
            Caption         =   ">"
            Height          =   255
            Left            =   2520
            TabIndex        =   77
            Top             =   1440
            Width           =   255
         End
         Begin VB.CommandButton Command3 
            Caption         =   "<"
            Height          =   255
            Left            =   2280
            TabIndex        =   76
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   3480
            Picture         =   "battle2.frx":14D831
            ScaleHeight     =   1095
            ScaleWidth      =   1095
            TabIndex        =   17
            Top             =   240
            Width           =   1095
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Left            =   1320
            TabIndex        =   15
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
         End
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   3480
            Picture         =   "battle2.frx":1536F3
            ScaleHeight     =   300
            ScaleWidth      =   345
            TabIndex        =   9
            Top             =   1440
            Width           =   345
         End
         Begin VB.Label Label18 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   66
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   42
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   41
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackColor       =   &H00000000&
            Caption         =   "Equip"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label9 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   2880
            TabIndex        =   16
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label5 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   3960
            TabIndex        =   10
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label4 
            BackColor       =   &H00000000&
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label3 
            BackColor       =   &H00000000&
            Caption         =   "Ability"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            Caption         =   "Transf."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            Caption         =   "Attack"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   0
         TabIndex        =   1
         Top             =   7320
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar2 
         Height          =   135
         Left            =   0
         TabIndex        =   2
         Top             =   7440
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar3 
         Height          =   135
         Left            =   0
         TabIndex        =   3
         Top             =   7560
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar5 
         Height          =   2655
         Left            =   5880
         TabIndex        =   21
         Top             =   1920
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   4683
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar6 
         Height          =   2655
         Left            =   3840
         TabIndex        =   79
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   4683
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   3195
         Left            =   4200
         Picture         =   "battle2.frx":153CD5
         ScaleHeight     =   3195
         ScaleWidth      =   1680
         TabIndex        =   18
         Top             =   1920
         Width           =   1680
         Begin VB.CommandButton Command1 
            Caption         =   "Change"
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   2760
            Width           =   1575
         End
      End
      Begin VB.Label Label28 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   1200
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label27 
         Caption         =   "0"
         Height          =   255
         Left            =   10560
         TabIndex        =   83
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label26 
         BackColor       =   &H00000000&
         Caption         =   "!Attacked!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   82
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         Caption         =   "!Attacked!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   81
         Top             =   4680
         Visible         =   0   'False
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dhp As Integer
Dim dlvl As Integer
Dim dmlvl As Integer
Dim dpwrpoints As Integer
Dim dpwr As Integer
Dim dmgc As Integer
Dim dmgcpoi As Integer
Dim darate As Integer
Dim dmrate As Integer
Dim dsrate As Integer

Dim dhead As Integer
Dim dbody As Integer
Dim drarm As Integer
Dim dlarm As Integer
Dim dtorso As Integer
Dim drleg As Integer
Dim dllef As Integer
Dim dweap As Integer

Dim ehead As Integer
Dim ebody As Integer
Dim erarm As Integer
Dim elarm As Integer
Dim etorso As Integer
Dim erleg As Integer
Dim ellef As Integer
Dim eweap As Integer

Dim ehp As Integer
Dim elvl As Integer
Dim emlvl As Integer
Dim epwr As Integer
Dim emgc As Integer
Dim earate As Integer
Dim emrate As Integer

Dim attackamount As Integer


Private Sub Command1_Click()
If Frame2.Visible = True Then
Frame2.Visible = False
Else
Frame2.Visible = True
End If

End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text + "<"
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text + ">"
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Load()
darate = 12
dmrate = 8
dsrate = 3

ProgressBar3.Value = 95

dlvl = 1
dmlvl = 2
dpwrpoints = 10
dpwr = 12
dmgc = 5
dmgcpoi = 5
darate = 12
dmrate = 8
dsrate = 5

dhead = 100
dbody = 100
drarm = 100
dlarm = 100
dtorso = 100
drleg = 100
dllef = 100
dweap = 100

ehead = 100
ebody = 100
erarm = 100
elarm = 100
etorso = 100
erleg = 100
ellef = 100
eweap = 100

elvl = 1
emlvl = 3
epwr = 15
emgc = 5
earate = 13
emrate = 7

checkbody
Module1.checkebody

End Sub

Private Sub Label1_Click()

If ProgressBar3.Value <= 20 Then
Label28.Caption = "Wait for your Stamina Bar! You ran out!"
Label28.Visible = True
End If

If Picture18.BorderStyle = 1 And ProgressBar3.Value > 20 Then
Label28.Visible = False
ehead = ehead - (ProgressBar1.Value / darate) - dpwr - (Time - 5) + 8
ProgressBar1.Value = 0
ProgressBar3.Value = ProgressBar3.Value - 15


End If

Label5.Caption = ehead
attackamount = attackamount + 1
checkebody
Timer1.Enabled = True
Timer6.Enabled = True
Label26.Visible = True
End Sub

Private Sub Label3_Click()
If Frame6.Visible = True Then
Frame6.Visible = False
Else
Frame6.Visible = True

End If
Frame4.Visible = False
Frame3.Visible = False
Frame5.Visible = False
End Sub

Private Sub Label4_Click()
If Frame4.Visible = True Then
Frame4.Visible = False
Else
Frame4.Visible = True

End If
Frame5.Visible = False
Frame3.Visible = False
Frame6.Visible = False
End Sub

Private Sub Label7_Click()
If Frame5.Visible = True Then
Frame5.Visible = False
Else
Frame5.Visible = True

End If
Frame4.Visible = False
Frame3.Visible = False
Frame6.Visible = False
End Sub

Private Sub List2_Click()

End Sub

Private Sub List2_ItemCheck(Item As Integer)

End Sub

Private Sub Picture18_Click()
Picture18.BorderStyle = 1
Picture21.BorderStyle = 0
Picture19.BorderStyle = 0
Picture20.BorderStyle = 0
Picture30.BorderStyle = 0
Picture29.BorderStyle = 0
Picture31.BorderStyle = 0
Picture34.BorderStyle = 0
End Sub

Private Sub Picture19_Click()
Picture19.BorderStyle = 1
Picture21.BorderStyle = 0
Picture18.BorderStyle = 0
Picture20.BorderStyle = 0
Picture30.BorderStyle = 0
Picture29.BorderStyle = 0
Picture31.BorderStyle = 0
Picture34.BorderStyle = 0
End Sub

Private Sub Picture20_Click()
Picture20.BorderStyle = 1
Picture21.BorderStyle = 0
Picture19.BorderStyle = 0
Picture18.BorderStyle = 0
Picture30.BorderStyle = 0
Picture29.BorderStyle = 0
Picture31.BorderStyle = 0
Picture34.BorderStyle = 0
End Sub

Private Sub Picture21_Click()
Picture21.BorderStyle = 1
Picture18.BorderStyle = 0
Picture19.BorderStyle = 0
Picture20.BorderStyle = 0
Picture30.BorderStyle = 0
Picture29.BorderStyle = 0
Picture31.BorderStyle = 0
Picture34.BorderStyle = 0
End Sub

Private Sub Picture29_Click()
Picture29.BorderStyle = 1
Picture21.BorderStyle = 0
Picture19.BorderStyle = 0
Picture20.BorderStyle = 0
Picture30.BorderStyle = 0
Picture18.BorderStyle = 0
Picture31.BorderStyle = 0
Picture34.BorderStyle = 0
End Sub

Private Sub Picture30_Click()
Picture30.BorderStyle = 1
Picture21.BorderStyle = 0
Picture19.BorderStyle = 0
Picture20.BorderStyle = 0
Picture18.BorderStyle = 0
Picture29.BorderStyle = 0
Picture31.BorderStyle = 0
Picture34.BorderStyle = 0
End Sub

Private Sub Picture31_Click()
Picture31.BorderStyle = 1
Picture21.BorderStyle = 0
Picture19.BorderStyle = 0
Picture20.BorderStyle = 0
Picture30.BorderStyle = 0
Picture29.BorderStyle = 0
Picture18.BorderStyle = 0
Picture34.BorderStyle = 0
End Sub

Private Sub Picture34_Click()
Picture34.BorderStyle = 1
Picture21.BorderStyle = 0
Picture19.BorderStyle = 0
Picture20.BorderStyle = 0
Picture30.BorderStyle = 0
Picture29.BorderStyle = 0
Picture31.BorderStyle = 0
Picture18.BorderStyle = 0
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value >= 88 Then
ProgressBar1.Value = 100
Timer1.Enabled = False
Else
ProgressBar1.Value = ProgressBar1.Value + darate


End If

If ProgressBar1.Value = 100 Then
Picture10.BackColor = &HFF&
End If

End Sub

Private Sub Timer4_Timer()
Label27.Caption = Label27.Caption + 1
End Sub

Private Sub Timer6_Timer()
If ProgressBar3.Value >= 92 Then
ProgressBar3.Value = 95
Timer6.Enabled = False
Else
ProgressBar3.Value = ProgressBar3.Value + dsrate


End If

If ProgressBar3.Value = 95 Then
Picture12.BackColor = &HFF&
End If
End Sub

Private Sub Timer8_Timer()
If ProgressBar2.Value >= 88 Then
ProgressBar2.Value = 100
Timer8.Enabled = False
Else
ProgressBar2.Value = ProgressBar2.Value + dmrate


End If

If ProgressBar2.Value = 100 Then
Picture11.BackColor = &HFF&
End If

End Sub

Private Sub Timer9_Timer()
If ProgressBar6.Value >= 90 Then
ProgressBar6.Value = 100
Else
ProgressBar6.Value = ProgressBar6.Value + 10

End If

If ProgressBar6.Value >= 80 Then

If eweap > 0 Then

If Label27.Caption < 13 Then
dlarm = dlarm - 23

End If


End If



End If
End Sub
Function checkbody()
If dhead >= 0 And dhead < 10 Then

Picture22.BackColor = &HFF&

ElseIf dhead >= 10 And dhead < 40 Then

Picture22.BackColor = &HFFFF&

ElseIf dhead >= 40 And dhead < 60 Then
Picture22.BackColor = &HC0FFFF

ElseIf dhead >= 60 And dhead < 80 Then
Picture22.BackColor = &HC0FFC0

ElseIf dhead >= 80 And dhead <= 100 Then
Picture22.BackColor = &HFF00&

End If
End Function
Function checkebody()
If ehead >= 0 And ehead < 10 Then

Picture18.BackColor = &HFF&

ElseIf ehead >= 10 And ehead < 40 Then

Picture18.BackColor = &HFFFF&

ElseIf ehead >= 40 And ehead < 60 Then
Picture18.BackColor = &HC0FFFF

ElseIf ehead >= 60 And ehead < 80 Then
Picture18.BackColor = &HC0FFC0

ElseIf ehead >= 80 And ehead <= 100 Then
Picture18.BackColor = &HFF00&

End If
End Function

