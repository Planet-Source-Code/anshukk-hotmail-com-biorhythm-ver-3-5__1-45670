VERSION 5.00
Begin VB.Form MainFrm 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Techno Rex BioRhythm Version 3.5"
   ClientHeight    =   7755
   ClientLeft      =   -885
   ClientTop       =   330
   ClientWidth     =   8520
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ComboBox Combo11 
      Height          =   315
      Left            =   960
      TabIndex        =   77
      Text            =   "Combo11"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000004&
      Height          =   7740
      Left            =   8640
      TabIndex        =   42
      Top             =   0
      Width           =   3100
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   53
         Top             =   3960
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   52
         Top             =   3480
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   51
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   47
         Top             =   2400
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Guide Me!!"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   88
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H00E0E0E0&
         Height          =   255
         Index           =   9
         Left            =   360
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "How To Compare two persons Biorhythm ?"
         Height          =   495
         Index           =   7
         Left            =   600
         TabIndex        =   50
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "How to Remove a Persons Name?"
         Height          =   495
         Index           =   6
         Left            =   600
         TabIndex        =   49
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "How To Add Persons Name?"
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   48
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "How To Get Today's Date in Condition Date Boxes?"
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   46
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "How To Check Biorhythm?"
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   44
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   2775
      End
      Begin VB.Image Image2 
         Height          =   7695
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3105
      End
   End
   Begin VB.ComboBox Combo15 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Main.frx":030A
      Left            =   960
      List            =   "Main.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox Combo14 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   600
      Width           =   885
   End
   Begin VB.ComboBox Combo13 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   765
   End
   Begin VB.ComboBox Combo12 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   765
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      Height          =   7740
      Left            =   8610
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   3100
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2040
         Width           =   885
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2040
         Width           =   885
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2040
         Width           =   885
      End
      Begin VB.ListBox List1 
         Height          =   3765
         ItemData        =   "Main.frx":030E
         Left            =   240
         List            =   "Main.frx":0310
         TabIndex        =   25
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Name"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   78
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Remove Name"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   89
         Top             =   2880
         Width           =   1995
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   255
         Index           =   11
         Left            =   600
         Top             =   2880
         Width           =   1995
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   255
         Index           =   10
         Left            =   600
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add/ Remove Names"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Default"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   7200
         Width           =   1695
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         Height          =   255
         Left            =   1080
         TabIndex        =   34
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   7080
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   7740
         Left            =   0
         Top             =   0
         Width           =   3120
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   7740
      Left            =   8610
      TabIndex        =   16
      Top             =   0
      Width           =   3100
      Begin VB.TextBox Text1 
         Height          =   5895
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         ItemData        =   "Main.frx":0312
         Left            =   120
         List            =   "Main.frx":0314
         TabIndex        =   18
         Text            =   "Combo7"
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Understanding Biorhythm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Question:"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   2055
      End
      Begin VB.Image Image4 
         Height          =   7740
         Left            =   0
         Top             =   0
         Width           =   3115
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   360
      MousePointer    =   2  'Cross
      ScaleHeight     =   4785
      ScaleWidth      =   7785
      TabIndex        =   15
      Top             =   1320
      Width           =   7815
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   7560
         TabIndex        =   90
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   7560
         TabIndex        =   93
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   7560
         TabIndex        =   92
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7080
         TabIndex        =   91
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   56
         Top             =   4560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Date and Time"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   55
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Label39"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   54
         Top             =   4560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   7440
         TabIndex        =   22
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "-100"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   21
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   20
         Top             =   0
         Width           =   375
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   30
         X1              =   7440
         X2              =   7440
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7800
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   0
         X1              =   240
         X2              =   240
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   1
         X1              =   480
         X2              =   480
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   2
         X1              =   720
         X2              =   720
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   3
         X1              =   960
         X2              =   960
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   4
         X1              =   1200
         X2              =   1200
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   5
         X1              =   1440
         X2              =   1440
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   6
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   7
         X1              =   1920
         X2              =   1920
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   8
         X1              =   2160
         X2              =   2160
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   9
         X1              =   2400
         X2              =   2400
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   10
         X1              =   2640
         X2              =   2640
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   11
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   12
         X1              =   3120
         X2              =   3120
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   13
         X1              =   3360
         X2              =   3360
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   14
         X1              =   3600
         X2              =   3600
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   3840
         X2              =   3840
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   16
         X1              =   4080
         X2              =   4080
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   17
         X1              =   4320
         X2              =   4320
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   18
         X1              =   4560
         X2              =   4560
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   19
         X1              =   4800
         X2              =   4800
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   20
         X1              =   5040
         X2              =   5040
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   21
         X1              =   5280
         X2              =   5280
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   22
         X1              =   5520
         X2              =   5520
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   23
         X1              =   5760
         X2              =   5760
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   24
         X1              =   6000
         X2              =   6000
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   25
         X1              =   6240
         X2              =   6240
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   26
         X1              =   6480
         X2              =   6480
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   27
         X1              =   6720
         X2              =   6720
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   28
         X1              =   6960
         X2              =   6960
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   29
         X1              =   7200
         X2              =   7200
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00400000&
         FillColor       =   &H00400000&
         FillStyle       =   0  'Solid
         Height          =   4335
         Left            =   3840
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Width           =   765
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   960
      Width           =   885
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   765
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   885
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Main.frx":0316
      Left            =   5280
      List            =   "Main.frx":0318
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   765
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Main.frx":031A
      Left            =   4440
      List            =   "Main.frx":031C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   765
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   80
      Top             =   7440
      Width           =   2535
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "                  Help               "
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   81
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Add /Remove Names      "
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   82
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Understanding Biorhythm "
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   83
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<<Today"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   84
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OFF"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7080
      TabIndex        =   79
      Top             =   600
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   8
      Left            =   7080
      Top             =   600
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   7
      Left            =   7080
      Top             =   960
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   255
      Index           =   6
      Left            =   5880
      Top             =   7440
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   255
      Index           =   5
      Left            =   5880
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   255
      Index           =   4
      Left            =   5880
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   255
      Index           =   3
      Left            =   5880
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Divine Cycle"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   3840
      TabIndex        =   85
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Primary Cycle"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   87
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Secondary Cycle"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   86
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   3840
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   2040
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   240
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      Height          =   255
      Left            =   8280
      TabIndex        =   96
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      Height          =   255
      Left            =   8280
      TabIndex        =   95
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      Height          =   255
      Left            =   8280
      TabIndex        =   94
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   59
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Emotional Cycle"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   60
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Spiritual Cycle"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   76
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   75
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Self Awareness Cycle"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3840
      TabIndex        =   74
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   73
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aesthetic  Cycle"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   72
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   71
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Intutive Cycle"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   70
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   69
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mastery Cycle"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   68
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   67
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Wisdom Cycle"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2040
      TabIndex        =   66
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   65
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Passion Cycle"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   64
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   63
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Intellectual Cycle"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   62
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Physical Cycle"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   58
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ">>Date Of Birth:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label34 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ">>Date Of Birth: "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   38
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Condition On Date:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   61
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   57
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Compare Mode"
      Height          =   615
      Left            =   7080
      TabIndex        =   97
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   7815
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim days As Integer
Dim spiritualCycleDays As Integer, physicalCycleDays As Integer, emotionalCycleDays As Integer, intellectualCycleDays As Integer, intutiveCycleDays As Integer, aestheticCycleDays As Integer, selfAwarenessCycleDays As Integer
Dim lastButtonPressed As Integer
Dim TopY As Double
Dim dateOfBirth As Date, dateOfCondition As Date, dateOfBirthCompare As Date
Dim bShow As Single
Dim CompareMode As Boolean
Dim GuideMode As Boolean
Dim GuideValue As Integer
Dim FromPaint As Boolean
Dim startupcode As String
Dim moveconstant As Single
Dim LastMoveEffect As Single
Dim GuideModePressed As Boolean
Dim GuideModePressedRemove As Boolean
Public language As String
Dim T(0 To 90) As String
' LANGUAGE VARIABLES
Dim OffButton As String
Dim OnButton As String
Dim m1  As String
Dim m2 As String
Dim HelpButton As String
Dim UBButton As String
Dim ARButton As String
Dim PercentLang As String

Dim WBiorhythm As String
Dim ABiorhythm As String
Dim WPhyCycle As String
Dim APhyCycle As String
Dim WEmoCycle As String
Dim AEmoCycle As String
Dim WIntCycle As String
Dim AIntCycle As String
Dim WPassCycle As String
Dim APassCycle As String
Dim WMastCycle As String
Dim AMastCycle As String
Dim WWisdCycle As String

Dim AWisdCycle As String
Dim WIntuCycle As String
Dim AIntuCycle As String
Dim WAesCycle As String
Dim AAesCycle As String
Dim WSelfCycle As String
Dim ASelfCycle As String
Dim WSpirCycle As String
Dim ASpirCycle As String
Dim WNP As String
Dim ANP As String
Dim WBal As String
Dim ABal As String

Dim S1 As String
Dim S2 As String
Dim S3 As String
Dim S4 As String
Dim S5 As String
Dim S6 As String
Dim S7 As String
Dim S8 As String
Dim S9 As String
Dim S10 As String
Dim S11 As String
Dim S12 As String
Dim S13 As String
Dim S14 As String



Private Sub Combo15_Click()
Dim strDate As String, strName As String
Dim syntaxFlag As Boolean, DoNotEnter As Boolean
Dim CntX As Integer, CntY As Integer

For CntX = 0 To List1.ListCount - 1
syntaxFlag = True
DoNotEnter = False

strName = ""
    For CntY = 1 To Len(List1.list(CntX))
        If "|" <> Mid((List1.list(CntX)), CntY, 1) And Not DoNotEnter Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                DoNotEnter = True
                If Combo15.Text = strName Then
                    If ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                           CntY = CntY + 1
                           syntaxFlag = False
                            
                    End If
                    If Not (syntaxFlag) Then
                            strDate = strDate & Mid((List1.list(CntX)), CntY, 1)
                    End If
                Else
                Exit For
                End If
                If Len(List1.list(CntX)) = CntY Then
                Combo12.Text = Day(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo13.Text = Month(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo14.Text = Year(CDate(Format(strDate, "dd/MM/yyyy")))
                End If
        End If

    Next
    
Next


            If GuideMode And (GuideValue = 52 Or GuideValue = 5) Then
            Call ShowGuideMeCallOut(S6, Label40(10))
            Else
            GuideMode = False
            End If
             
            If GuideValue <> 52 Or GuideValue <> 5 Then
            GuideMode = False
            End If


End Sub




Private Sub Combo11_Click()
Dim strDate As String, strName As String
Dim syntaxFlag As Boolean, DoNotEnter As Boolean
Dim CntX As Integer, CntY As Integer

For CntX = 0 To List1.ListCount - 1
syntaxFlag = True
DoNotEnter = False

strName = ""
    For CntY = 1 To Len(List1.list(CntX))
        If "|" <> Mid((List1.list(CntX)), CntY, 1) And Not DoNotEnter Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                DoNotEnter = True
                If Combo11.Text = strName Then
                    If ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                           CntY = CntY + 1
                           syntaxFlag = False
                            
                    End If
                    If Not (syntaxFlag) Then
                            strDate = strDate & Mid((List1.list(CntX)), CntY, 1)
                    End If
                Else
                Exit For
                End If
                If Len(List1.list(CntX)) = CntY Then
                Combo1.Text = Day(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo2.Text = Month(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo3.Text = Year(CDate(Format(strDate, "dd/MM/yyyy")))
                End If
        End If

    Next
    
Next

'THE GUIDE ME SECTION
If GuideMode Then
    If CompareMode Then
                    If GuideValue = 5 Then
                            Call ShowGuideMeCallOut(S7, Combo15)
                    End If
                            
                    If GuideValue = 1 Then
                    Call ShowGuideMeCallOut(S8, Label40(0))
                            GuideValue = 12
                    End If
            
            Else
                    If GuideValue = 1 Then
                    Call ShowGuideMeCallOut(S9, Label40(10))
                    End If
                    
                    If GuideValue = 5 Then
                    Call ShowGuideMeCallOut(S10, Label40(0))
                    GuideValue = 52
                    End If
                    
    End If
End If

If (GuideValue <> 5 And GuideValue <> 1) And (GuideValue <> 52 And GuideValue <> 12) Then
GuideMode = False
End If


End Sub

'MAKE CURVES FORWARDS
 Sub MakeCurve(x As Double, y As Double, d As Double, c As Long)
 Dim T
 Picture1.ForeColor = c
 Picture1.PSet (x, y)
 For T = x To days * 31
    Picture1.Line -(T, y - (2160 * Sin((3.14) * ((T - x) / ((d - x) / 2)))))
 Next
End Sub


'MAKES CURVES BACKWARDS
Sub MakeReverseCurve(x As Double, y As Double, d As Double, c As Long)
Dim T
Picture1.ForeColor = c
Picture1.PSet (days - d + x, y)
For T = (days - d + x) To x
    Picture1.Line -(T, y - (2160 * Sin((3.14) * ((T - x) / ((d - x) / 2)))))
Next
End Sub

'MAKES SECONDARY CURVE
Sub MakeSecCurve(x1 As Double, y1 As Double, d1 As Double, c1 As Long, x2 As Double, y2 As Double, D2 As Double)
Dim T, x
 Picture1.ForeColor = c1
 Picture1.PSet (0, ((y1 - (2160 * Sin((3.14) * ((T - x1) / ((d1 - x1) / 2))))) + (y2 - (2160 * Sin((3.14) * ((T - x2) / ((D2 - x2) / 2)))))) / 2)
 For T = x To days * 31
    Picture1.Line -(T, ((y1 - (2160 * Sin((3.14) * ((T - x1) / ((d1 - x1) / 2))))) + (y2 - (2160 * Sin((3.14) * ((T - x2) / ((D2 - x2) / 2)))))) / 2)
 Next
End Sub

' CHECKS WHETHER YEAR IS LEAP YEAR OR NOT
Function bIsLeapYear(ByVal inYear As Integer) As Boolean
    bIsLeapYear = ((inYear Mod 4 = 0) _
               And (inYear Mod 100 <> 0) _
                Or (inYear Mod 400 = 0))
End Function
Private Sub Combo7_click()
'PUTS HELP TEXT IN TEXT BOX  WHEN COMBO7 IS CLICKED
Select Case Combo7.Text
Case WBiorhythm
Text1.Text = ABiorhythm
Case WPhyCycle
Text1.Text = APhyCycle
Case WEmoCycle
Text1.Text = AEmoCycle
Case WIntCycle
Text1.Text = AIntCycle
Case WPassCycle
Text1.Text = APassCycle
Case WMastCycle
Text1.Text = AMastCycle
Case WWisdCycle
Text1.Text = AWisdCycle

Case WIntuCycle
Text1.Text = AIntuCycle

Case WAesCycle
Text1.Text = AAesCycle
Case WSelfCycle
Text1.Text = ASelfCycle
Case WSpirCycle
Text1.Text = ASpirCycle

Case WNP
Text1.Text = ANP
Case WBal
Text1.Text = ABal
End Select
End Sub



Private Sub ComcmdHelp_Click()
Load Form1
Form1.Show
End Sub




Private Sub Combo7_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub Form_Load()
On Error Resume Next
Dim CntX As Integer, CntY As Integer
Dim strName As String, strDate As String
Dim syntaxFlag As Boolean, DoNotEnter As Boolean
'SET THE SKIN FILE
If FileExists(App.Path & "\skin.jpg") Then
Image1.Picture = LoadPicture(App.Path & "\SKIN.JPG")
Else
MsgBox "It seems that you donot have any Skin file." & vbCrLf & "Please download the Skin File(s) from " & vbCrLf & vbCrLf & "http://www.geocities.com/anshukk", vbInformation
End If

'SET THE LANGUAGE
SetLanguage (language)
Combo7.Text = WBiorhythm
Combo7_click

'PUTING QUESTIONS IN UNDERSTANDING BIORHYTHM AND OTHER SETTING
bShow = 0
Combo7.AddItem WBiorhythm
Combo7.AddItem WPhyCycle
Combo7.AddItem WEmoCycle
Combo7.AddItem WIntCycle
Combo7.AddItem WPassCycle
Combo7.AddItem WMastCycle
Combo7.AddItem WWisdCycle
Combo7.AddItem WIntuCycle
Combo7.AddItem WAesCycle
Combo7.AddItem WSelfCycle
Combo7.AddItem WSpirCycle
Combo7.AddItem WNP
Combo7.AddItem WBal


Call OptimizeFormPosition


'SETS THE NUMBER OF DAYS FOR THE VARIOUS CYCLES

physicalCycleDays = 23
emotionalCycleDays = 28
intellectualCycleDays = 33
intutiveCycleDays = 38
aestheticCycleDays = 43
selfAwarenessCycleDays = 48
spiritualCycleDays = 53

'SETS VALUES OF DAYS AND POSITION OF THE WAVE
days = 240
TopY = 2280


'PUTS YEARS IN THE COMBOBOX
For CntX = 1900 To 2010
Combo3.AddItem CntX
Combo6.AddItem CntX
Combo10.AddItem CntX
Combo14.AddItem CntX
Next

'PUTS MONTHS IN COMBOBOX

For CntX = 1 To 12
Combo2.AddItem CntX
Combo5.AddItem CntX
Combo9.AddItem CntX
Combo13.AddItem CntX
Next

'PUTS DATES IN COMBOBOX

For CntX = 1 To 31
Combo1.AddItem CntX
Combo4.AddItem CntX
Combo8.AddItem CntX
Combo12.AddItem CntX
Next

'PUTS TODAY'S DATE IN CONDITION DATE COMBOBOX

Combo5.Text = Month(Date)
Combo6.Text = Year(Date)
Combo4.Text = Day(Date)

'SETS COLOR FOR LEGEND
Label14.BackColor = RGB(255, 200, 100)
Label16.BackColor = RGB(100, 200, 255)
Label18.BackColor = RGB(200, 255, 100)



'LOADING SAVED NAMES AND DOB'S
Call List_Load(List1, App.Path & "\DOB_BRHTHM35.DAT")

'LOAD NAMES IN COMBOBOX 11

For CntX = 0 To List1.ListCount - 1
    For CntY = 1 To Len(List1.list(CntX))
        If Not ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                Combo11.AddItem strName
                strName = ""
                Exit For
               
        End If
    Next
Next


'LOAD NAMES IN COMBOBOX 15



For CntX = 0 To List1.ListCount - 1
    For CntY = 1 To Len(List1.list(CntX))
        If Not ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                Combo15.AddItem strName
                strName = ""
                Exit For
               
        End If
    Next
Next

'LOADS DEFAULT NAME IN NAME COMBOBOX
If List1.ListCount <> 0 Then
CntX = 0
syntaxFlag = True
DoNotEnter = False

strName = ""
    For CntY = 1 To Len(List1.list(CntX))
        If "|" <> Mid((List1.list(CntX)), CntY, 1) And Not DoNotEnter Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                DoNotEnter = True
                    If ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                           CntY = CntY + 1
                           syntaxFlag = False
                            
                    End If
                    If Not (syntaxFlag) Then
                            strDate = strDate & Mid((List1.list(CntX)), CntY, 1)
                    End If
                If Len(List1.list(CntX)) = CntY Then
                Combo1.Text = Day(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo2.Text = Month(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo3.Text = Year(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo11.Text = strName
                Combo12.Text = Day(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo13.Text = Month(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo14.Text = Year(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo15.Text = strName
                End If
        End If
    Next
End If


'SETTING INITIAL VALUE OF COMPAREMODE GUIDEMODE AND GUIDEVALUE
CompareMode = False
GuideMode = False
GuideValue = 0
FromPaint = False
GuideModePressed = True
GuideModePressedRemove = True

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call MoveEffect(Label40, 32, 11)
End Sub

Private Sub Form_Paint()
' RESET OR REPAINTS THE GRAPH IF THE FORM HAS BEEN MINIMIZED
' OR PART OR ALL OF AN OBJECT IS EXPOSED AFTER BEING MOVED
' OR ENLARGED,
' OR AFTER A WINDOW THAT WAS COVERING THE OBJECT HAS BEEN MOVED
FromPaint = True
If lastButtonPressed = 1 Then
Label40_Click (11)
Exit Sub
End If

If lastButtonPressed = 2 Then
Label40_Click (10)
Exit Sub
End If

If lastButtonPressed = 3 Then
Label40_Click (9)
Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call List_Save(List1, App.Path & "\DOB_BRHTHM35.DAT")
End Sub





Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call MoveEffect(Label40, 32, 11)
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call MoveEffect(Label40, 32, 11)
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call MoveEffect(Label40, 32, 11)
End Sub

Private Sub Label40_Click(Index As Integer)
'DECLARATION FOR 1 AND 2  ==================
Dim CntY As Integer
Dim CntX As Integer
Dim strName As String, strDate As String
Dim DoNotEnter As Boolean, syntaxFlag As Boolean
'DECLARATION FOR 11 , 10 AND 9 =============
Dim ndays As Double
Dim b As Integer, r As Integer, g As Integer, w As Integer
Dim l As Double, m As Double, n As Double, o As Double
Dim TopYCompare As Double
Dim bc, rc, gc, oc As Integer
Dim lc, mc, nc, wc As Double
Dim ndaysCompare As Double
Dim TopYCompare1 As Double
Dim TopYCompare2 As Double
'===========================================

Select Case (Index)
'=========================================================
'=========================================================
'CASE 0 COMPARE MODE ON/OFF SWITCH
'=========================================================
'=========================================================

Case 0 ' COMPARE SWITCH
Label4.Caption = ""
Label6.Caption = ""
Label8.Caption = ""

Label10.Caption = ""
Label14.Caption = ""
Label16.Caption = ""
Label18.Caption = ""

Label3(0).Caption = ""
Label3(2).Caption = ""
Label3(4).Caption = ""

If CompareMode Then
Combo12.Enabled = False
Combo13.Enabled = False
Combo14.Enabled = False
Combo15.Enabled = False

Combo12.BackColor = &HE0E0E0
Combo13.BackColor = &HE0E0E0
Combo14.BackColor = &HE0E0E0
Combo15.BackColor = &HE0E0E0


Label40(0).Caption = OffButton
CompareMode = False
Else
Combo12.Enabled = True
Combo13.Enabled = True
Combo14.Enabled = True
Combo15.Enabled = True

Combo12.BackColor = &H80000009
Combo13.BackColor = &H80000009
Combo14.BackColor = &H80000009
Combo15.BackColor = &H80000009

Label40(0).Caption = OnButton
CompareMode = True
End If

'THE GUIDE ME SECTION
If GuideMode Then
If GuideValue = 52 Then
Call ShowGuideMeCallOut(S11, Combo15)
End If

If GuideValue = 12 Then
        Call ShowGuideMeCallOut(S12, Label40(10))
End If
End If
        
If GuideValue <> 52 And GuideValue <> 12 Then
GuideMode = False
End If

'=========================================================
'=========================================================
'CASE 1 REMOVE THE NAME FROM DATABASE
'=========================================================
'=========================================================


Case 1 ' REMOVE NAME


GuideMode = False

'REMOVES NAME FROM LIST
If List1.ListIndex = -1 Then
MsgBox m1
Else
List1.RemoveItem List1.ListIndex
End If

'REFRESH NAMES IN COMBOBOX
Combo11.Clear
For CntX = 0 To List1.ListCount - 1
    For CntY = 1 To Len(List1.list(CntX))
        If Not ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                Combo11.AddItem strName
                strName = ""
                Exit For
                
        End If
    Next
Next



'REFRESH THE NAME COMPARISION COMBOBOX
Combo15.Clear
For CntX = 0 To List1.ListCount - 1
    For CntY = 1 To Len(List1.list(CntX))
        If Not ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                Combo15.AddItem strName
                strName = ""
                Exit For
                
        End If
    Next
Next


'LOADS DEFAULT NAME IN NAME COMBOBOX
If List1.ListCount <> 0 Then
CntX = 0
syntaxFlag = True
DoNotEnter = False

strName = ""
    For CntY = 1 To Len(List1.list(CntX))
        If "|" <> Mid((List1.list(CntX)), CntY, 1) And Not DoNotEnter Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                DoNotEnter = True
                    If ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                           CntY = CntY + 1
                           syntaxFlag = False
                            
                    End If
                    If Not (syntaxFlag) Then
                            strDate = strDate & Mid((List1.list(CntX)), CntY, 1)
                    End If
                If Len(List1.list(CntX)) = CntY Then
                Combo1.Text = Day(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo2.Text = Month(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo3.Text = Year(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo11.Text = strName
                End If
        End If
    Next
End If
'=========================================================
'=========================================================
'CASE 2 ADDS NAMES TO THE DATABASE
'=========================================================
'=========================================================


Case 2

GuideMode = False

'CHECK VALIDITY OF DATE
If Not (IsDate(Combo8.Text & "/" & Combo9.Text & "/" & Combo10.Text)) Then
    MsgBox m2
    Exit Sub
End If

'CHECKS FOR BLANK NAMES AND ADDS NAME TO LIST
If LTrim(RTrim(Text2.Text)) <> "" Then
List1.AddItem Text2.Text & "|" & Combo8.Text & "/" & Combo9.Text & "/" & Combo10.Text

'REFRESH THE NAME COMBOBOX
Combo11.Clear
For CntX = 0 To List1.ListCount - 1
    For CntY = 1 To Len(List1.list(CntX))
        If Not ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                Combo11.AddItem strName
                strName = ""
                Exit For
                
        End If
    Next
Next

'REFRESH THE NAME COMPARISION COMBOBOX
Combo15.Clear
For CntX = 0 To List1.ListCount - 1
    For CntY = 1 To Len(List1.list(CntX))
        If Not ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                Combo15.AddItem strName
                strName = ""
                Exit For
                
        End If
    Next
Next


'LOADS DEFAULT NAME IN NAME COMBOBOX
If List1.ListCount <> 0 Then
CntX = 0
syntaxFlag = True
DoNotEnter = False

strName = ""
    For CntY = 1 To Len(List1.list(CntX))
        If "|" <> Mid((List1.list(CntX)), CntY, 1) And Not DoNotEnter Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                DoNotEnter = True
                    If ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                           CntY = CntY + 1
                           syntaxFlag = False
                            
                    End If
                    If Not (syntaxFlag) Then
                            strDate = strDate & Mid((List1.list(CntX)), CntY, 1)
                    End If
                If Len(List1.list(CntX)) = CntY Then
                Combo1.Text = Day(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo2.Text = Month(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo3.Text = Year(CDate(Format(strDate, "dd/MM/yyyy")))
                Combo11.Text = strName
                End If
        End If
    Next
End If
Else
MsgBox "Please enter Name"
End If
'=========================================================
'=========================================================
'CASE 3 EXIT
'=========================================================
'=========================================================


Case 3 'EXIT BUTTON
Unload Me
'=========================================================
'=========================================================
'CASE 4 OPENS THE HELP SIDE WINDOW
'=========================================================
'=========================================================

Case 4

Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Set Image2.Picture = Image1.Picture

If Label46.Caption = "<<" Then
Label46.Caption = ">>"
Else
Label44.Caption = ">>"
Label45.Caption = ">>"
Label46.Caption = "<<"
Me.Width = 11820
End If


If Label44.Caption = ">>" And _
Label45.Caption = ">>" And _
Label46.Caption = ">>" Then
Me.Width = 8610
End If

Call OptimizeFormPosition
'=========================================================
'=========================================================
'CASE 5 OPENS THE ADD/REMOVE NAME SIDE WINDOW
'=========================================================
'=========================================================


Case 5
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Set Image3.Picture = Image1.Picture

If Label45.Caption = "<<" Then
Label45.Caption = ">>"
Else
Label44.Caption = ">>"
Label45.Caption = "<<"
Label46.Caption = ">>"
Me.Width = 11820
End If



If Label44.Caption = ">>" And _
Label45.Caption = ">>" And _
Label46.Caption = ">>" Then
Me.Width = 8610
End If

Call OptimizeFormPosition

'THE GUIDE ME SECTION

If GuideMode And GuideValue = 3 And GuideModePressed Then
Call ShowGuideMeCallOut(S13, Label40(2))
GuideModePressed = False
End If

If GuideMode And GuideValue = 4 And GuideModePressedRemove Then
Call ShowGuideMeCallOut(S14, Label40(1))
GuideModePressedRemove = False
End If
'=========================================================
'=========================================================
'CASE 6 SHOWS AND HIDES THE UNDERSTANDING BIORHYTHM FRAME
'=========================================================
'=========================================================

Case 6
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Set Image4.Picture = Image1.Picture

If Label44.Caption = "<<" Then
Label44.Caption = ">>"
Else
Label44.Caption = "<<"
Label45.Caption = ">>"
Label46.Caption = ">>"
Me.Width = 11820
End If


If Label44.Caption = ">>" And _
Label45.Caption = ">>" And _
Label46.Caption = ">>" Then
Me.Width = 8610
End If

Call OptimizeFormPosition
'=========================================================
'=========================================================
'CASE 7 PUTS CURRENT DATE IN CONDITION DATE COMBO BOXS
'=========================================================
'=========================================================


Case 7
Combo5.Text = Month(Date)
Combo6.Text = Year(Date)
Combo4.Text = Day(Date)

'=========================================================
'=========================================================
'CASE 8 GUIDE ME!!! BUTTON
'=========================================================
'=========================================================



Case 8
Dim Index1 As Integer

For Index1 = 1 To 5
If Option1(Index1).Value = True Then
Label40_Click (4)
GuideValue = Index1
GuideMode = True
Exit For
End If
Next
Select Case (GuideValue)
Case 1

Call ShowGuideMeCallOut(S1, Combo11)
GuideModePressed = True
GuideModePressedRemove = True
Case 2
Call ShowGuideMeCallOut(S2, Label40(7))
GuideMode = False
GuideModePressed = True
GuideModePressedRemove = True
Case 3
Call ShowGuideMeCallOut(S3, Label40(5))
GuideModePressed = True
GuideModePressedRemove = True
Case 4
Call ShowGuideMeCallOut(S4, Label40(5))
GuideModePressed = True
GuideModePressedRemove = True
Case 5
Call ShowGuideMeCallOut(S5, Combo11)
GuideModePressed = True
GuideModePressedRemove = True
End Select

'=========================================================
'=========================================================
'CASE 9 DIVINE CYCLE
'=========================================================
'=========================================================

Case 9


Label3(0).Caption = ""
Label3(2).Caption = ""
Label3(4).Caption = ""

Label4.Caption = ""
Label6.Caption = ""
Label8.Caption = ""

If Not FromPaint Then
GuideMode = False
End If
FromPaint = False


If Not CompareMode Then
'CHECKS FOR VALIDITY OF DATE
If Not (IsDate(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text) And IsDate(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text)) Then
    MsgBox "The Entered date is not valid."
    Exit Sub
End If

Picture1.Cls
lastButtonPressed = 3
Label20.Visible = True

dateOfBirth = Format(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text, "dd/MM/yyyy")
dateOfCondition = Format(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text, "dd/MM/yyyy")


'FINDS THE AGE OF PERSON IN DAYS ON CONDITION DATE
ndays = DateDiff("d", dateOfBirth, dateOfCondition)

'FINDS THE EXACT POSITION OF EACH CYCLE
l = ((ndays / intutiveCycleDays - Int(ndays / intutiveCycleDays)) * intutiveCycleDays)
m = ((ndays / aestheticCycleDays - Int(ndays / aestheticCycleDays)) * aestheticCycleDays)
n = ((ndays / selfAwarenessCycleDays - Int(ndays / selfAwarenessCycleDays)) * selfAwarenessCycleDays)
o = ((ndays / spiritualCycleDays - Int(ndays / spiritualCycleDays)) * spiritualCycleDays)

b = (16 - l) * days
r = (16 - m) * days
g = (16 - n) * days
w = (16 - o) * days



'DRAWS WAVES FROM CENTER TO BACKWARD AND CENTER TO FORWARD
Call MakeCurve(days + b, TopY, (days * intutiveCycleDays) + b, RGB(255, 200, 100))
Call MakeReverseCurve(days + b, TopY, (days * intutiveCycleDays) + b, RGB(255, 200, 100))

Call MakeCurve(days + r, TopY, (days * aestheticCycleDays) + r, RGB(100, 200, 255))
Call MakeReverseCurve(days + r, TopY, (days * aestheticCycleDays) + r, RGB(100, 200, 255))

Call MakeCurve(days + g, TopY, (days * selfAwarenessCycleDays) + g, RGB(200, 255, 100))
Call MakeReverseCurve(days + g, TopY, (days * selfAwarenessCycleDays) + g, RGB(200, 255, 100))

Call MakeCurve(days + w, TopY, (days * spiritualCycleDays) + w, RGB(255, 255, 255))
Call MakeReverseCurve(days + w, TopY, (days * spiritualCycleDays) + w, RGB(255, 255, 255))

Else 'COMPARE MODE

'CHECKS FOR VALIDITY OF DATE
If Not (IsDate(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text) And IsDate(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text)) Then
    MsgBox "The Entered date is not valid."
    Exit Sub
End If

Picture1.Cls
lastButtonPressed = 3
Label20.Visible = True

dateOfBirth = Format(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text, "dd/MM/yyyy")
dateOfCondition = Format(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text, "dd/MM/yyyy")
dateOfBirthCompare = Format(Combo12.Text & "/" & Combo13.Text & "/" & Combo14.Text, "dd/mm/yyyy")

'FINDS THE AGE OF PERSON IN DAYS ON CONDITION DATE
ndays = DateDiff("d", dateOfBirth, dateOfCondition)

'FINDS THE EXACT POSITION OF EACH CYCLE
l = ((ndays / intutiveCycleDays - Int(ndays / intutiveCycleDays)) * intutiveCycleDays)
m = ((ndays / aestheticCycleDays - Int(ndays / aestheticCycleDays)) * aestheticCycleDays)
n = ((ndays / selfAwarenessCycleDays - Int(ndays / selfAwarenessCycleDays)) * selfAwarenessCycleDays)
o = ((ndays / spiritualCycleDays - Int(ndays / spiritualCycleDays)) * spiritualCycleDays)

b = (16 - l) * days
r = (16 - m) * days
g = (16 - n) * days
w = (16 - o) * days

'FINDS THE AGE OF PERSON IN DAYS ON CONDITION DATE
ndaysCompare = DateDiff("d", dateOfBirthCompare, dateOfCondition)

'FINDS THE EXACT POSITION OF EACH CYCLE
lc = ((ndaysCompare / intutiveCycleDays - Int(ndaysCompare / intutiveCycleDays)) * intutiveCycleDays)
mc = ((ndaysCompare / aestheticCycleDays - Int(ndaysCompare / aestheticCycleDays)) * aestheticCycleDays)
nc = ((ndaysCompare / selfAwarenessCycleDays - Int(ndaysCompare / selfAwarenessCycleDays)) * selfAwarenessCycleDays)
oc = ((ndaysCompare / spiritualCycleDays - Int(ndaysCompare / spiritualCycleDays)) * spiritualCycleDays)

bc = (16 - lc) * days
rc = (16 - mc) * days
gc = (16 - nc) * days
wc = (16 - oc) * days

If dateOfBirthCompare > dateOfBirth Then

'DRAWS WAVES FROM CENTER TO BACKWARD AND CENTER TO FORWARD
Label14.Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intutiveCycleDays), 0) & "%"
TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intutiveCycleDays)) + TopY
Call MakeCurveCompare(days + bc, TopYCompare, (days * intutiveCycleDays) / 2 + bc, RGB(255, 200, 100))
Call MakeReverseCurveCompare(days + bc, TopYCompare, (days * intutiveCycleDays) / 2 + bc, RGB(255, 200, 100))

Label16.Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, aestheticCycleDays), 0) & "%"
TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, aestheticCycleDays)) + TopY
Call MakeCurveCompare(days + rc, TopYCompare, (days * aestheticCycleDays) / 2 + rc, RGB(100, 200, 255))
Call MakeReverseCurveCompare(days + rc, TopYCompare, (days * aestheticCycleDays) / 2 + rc, RGB(100, 200, 255))

Label18.Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, selfAwarenessCycleDays), 0) & "%"
TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, selfAwarenessCycleDays)) + TopY
Call MakeCurveCompare(days + gc, TopYCompare, (days * selfAwarenessCycleDays) / 2 + gc, RGB(200, 255, 100))
Call MakeReverseCurveCompare(days + gc, TopYCompare, (days * selfAwarenessCycleDays) / 2 + gc, RGB(200, 255, 100))

Label10.Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, spiritualCycleDays), 0) & "%"
TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, spiritualCycleDays)) + TopY
Call MakeCurveCompare(days + wc, TopYCompare, (days * spiritualCycleDays) / 2 + wc, RGB(255, 255, 255))
Call MakeReverseCurveCompare(days + wc, TopYCompare, (days * spiritualCycleDays) / 2 + wc, RGB(255, 255, 255))
Else

'DRAWS WAVES FROM CENTER TO BACKWARD AND CENTER TO FORWARD
Label14.Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intutiveCycleDays), 0) & "%"
TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intutiveCycleDays)) + TopY
Call MakeCurveCompare(days + b, TopYCompare, (days * intutiveCycleDays) / 2 + b, RGB(255, 200, 100))
Call MakeReverseCurveCompare(days + b, TopYCompare, (days * intutiveCycleDays) / 2 + b, RGB(255, 200, 100))

Label16.Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, aestheticCycleDays), 0) & "%"
TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, aestheticCycleDays)) + TopY
Call MakeCurveCompare(days + r, TopYCompare, (days * aestheticCycleDays) / 2 + r, RGB(100, 200, 255))
Call MakeReverseCurveCompare(days + r, TopYCompare, (days * aestheticCycleDays) / 2 + r, RGB(100, 200, 255))

Label18.Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, selfAwarenessCycleDays), 0) & "%"
TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, selfAwarenessCycleDays)) + TopY
Call MakeCurveCompare(days + g, TopYCompare, (days * selfAwarenessCycleDays) / 2 + g, RGB(200, 255, 100))
Call MakeReverseCurveCompare(days + g, TopYCompare, (days * selfAwarenessCycleDays) / 2 + g, RGB(200, 255, 100))

Label10.Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, spiritualCycleDays), 0) & "%"
TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, spiritualCycleDays)) + TopY
Call MakeCurveCompare(days + w, TopYCompare, (days * spiritualCycleDays) / 2 + w, RGB(255, 255, 255))
Call MakeReverseCurveCompare(days + w, TopYCompare, (days * spiritualCycleDays) / 2 + w, RGB(255, 255, 255))
End If


End If
Label20.Visible = True
Label39.Visible = True
Label21.Visible = True

'=========================================================
'=========================================================
'CASE 10 SECONDARY CYCLE
'=========================================================
'=========================================================


Case 10


Label10.Caption = ""
Label14.Caption = ""
Label16.Caption = ""
Label18.Caption = ""

Label3(0).Caption = ""
Label3(2).Caption = ""
Label3(4).Caption = ""

If Not FromPaint Then
GuideMode = False
End If
FromPaint = False


If Not CompareMode Then

'CHECKS FOR VALIDITY OF DATE

If Not (IsDate(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text) And IsDate(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text)) Then
MsgBox "The Entered date is not valid."
Exit Sub
End If

Picture1.Cls
lastButtonPressed = 2
Label20.Visible = True


dateOfBirth = Format(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text, "dd/MM/yyyy")
dateOfCondition = Format(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text, "dd/MM/yyyy")

'FINDS THE AGE OF PERSON IN DAYS ON CONDITION DATE

ndays = DateDiff("d", dateOfBirth, dateOfCondition)
'FINDS THE EXACT POSITION OF EACH CYCLE

l = (ndays / physicalCycleDays - Int(ndays / physicalCycleDays)) * physicalCycleDays
m = (ndays / emotionalCycleDays - Int(ndays / emotionalCycleDays)) * emotionalCycleDays
n = (ndays / intellectualCycleDays - Int(ndays / intellectualCycleDays)) * intellectualCycleDays
 
b = (16 - l) * days
r = (16 - m) * days
g = (16 - n) * days

'MAKES SECONDARY WAVE WHICH IS COMBINATION OF TWO PRIMARY WAVES
Call MakeSecCurve(days + b, TopY, (days * physicalCycleDays) + b, RGB(255, 255, 0), days + r, TopY, (days * emotionalCycleDays) + r)
Call MakeSecCurve(days + g, TopY, (days * intellectualCycleDays) + g, RGB(255, 0, 255), days + r, TopY, (days * emotionalCycleDays) + r)
Call MakeSecCurve(days + g, TopY, (days * intellectualCycleDays) + g, RGB(0, 255, 255), days + b, TopY, (days * physicalCycleDays) + b)
Else 'COMPARE MODE

Picture1.Cls
lastButtonPressed = 2
Label20.Visible = True

dateOfBirthCompare = Format(Combo12.Text & "/" & Combo13.Text & "/" & Combo14.Text, "dd/mm/yyyy")
dateOfBirth = Format(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text, "dd/MM/yyyy")
dateOfCondition = Format(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text, "dd/MM/yyyy")

'FINDS THE AGE OF PERSON IN DAYS ON CONDITION DATE

ndays = DateDiff("d", dateOfBirth, dateOfCondition)
'FINDS THE EXACT POSITION OF EACH CYCLE

l = (ndays / physicalCycleDays - Int(ndays / physicalCycleDays)) * physicalCycleDays
m = (ndays / emotionalCycleDays - Int(ndays / emotionalCycleDays)) * emotionalCycleDays
n = (ndays / intellectualCycleDays - Int(ndays / intellectualCycleDays)) * intellectualCycleDays
 
b = (16 - l) * days
r = (16 - m) * days
g = (16 - n) * days

'FINDS THE AGE OF PERSON IN DAYS ON CONDITION DATE

ndaysCompare = DateDiff("d", dateOfBirth, dateOfBirthCompare)
'FINDS THE EXACT POSITION OF EACH CYCLE

lc = (ndaysCompare / physicalCycleDays - Int(ndaysCompare / physicalCycleDays)) * physicalCycleDays
mc = (ndaysCompare / emotionalCycleDays - Int(ndaysCompare / emotionalCycleDays)) * emotionalCycleDays
nc = (ndaysCompare / intellectualCycleDays - Int(ndaysCompare / intellectualCycleDays)) * intellectualCycleDays
 
bc = (16 - lc) * days
rc = (16 - mc) * days
gc = (16 - nc) * days
TopYCompare1 = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, physicalCycleDays)) + TopY
TopYCompare2 = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, emotionalCycleDays)) + TopY
Label4.Caption = Round((GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, physicalCycleDays) + GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, emotionalCycleDays)) / 2, 0) & "%"
Call MakeSecCurveCompare(days + b, days + r, days + bc, days + rc, TopYCompare1, TopYCompare2, TopYCompare1, TopYCompare2, (days * physicalCycleDays), days * emotionalCycleDays, days * physicalCycleDays, days * emotionalCycleDays, RGB(255, 255, 0))

TopYCompare1 = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intellectualCycleDays)) + TopY
TopYCompare2 = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, emotionalCycleDays)) + TopY
Label6.Caption = Round((GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intellectualCycleDays) + GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, emotionalCycleDays)) / 2, 0) & "%"
Call MakeSecCurveCompare(days + g, days + r, days + gc, days + rc, TopYCompare1, TopYCompare2, TopYCompare1, TopYCompare2, (days * intellectualCycleDays), days * emotionalCycleDays, days * intellectualCycleDays, days * emotionalCycleDays, RGB(255, 0, 255))

TopYCompare1 = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, physicalCycleDays)) + TopY
TopYCompare2 = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intellectualCycleDays)) + TopY
Label8.Caption = Round((GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, physicalCycleDays) + GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intellectualCycleDays)) / 2, 0) & "%"
Call MakeSecCurveCompare(days + b, days + g, days + bc, days + gc, TopYCompare1, TopYCompare2, TopYCompare1, TopYCompare2, (days * physicalCycleDays), days * intellectualCycleDays, days * physicalCycleDays, days * intellectualCycleDays, RGB(0, 255, 255))


End If

Label20.Visible = True
Label39.Visible = True
Label21.Visible = True



'=========================================================
'=========================================================
'CASE 11 'PRIMARY CYCLE
'=========================================================
'=========================================================

Case 11


Label4.Caption = ""
Label6.Caption = ""
Label8.Caption = ""

Label10.Caption = ""
Label14.Caption = ""
Label16.Caption = ""
Label18.Caption = ""


If Not FromPaint Then
GuideMode = False
End If
FromPaint = False



If Not CompareMode Then
        'CHECKS FOR VALIDITY OF DATE
        
        If Not (IsDate(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text) And IsDate(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text)) Then
            MsgBox "The Entered date is not valid."
            Exit Sub
        End If
        
        
        Picture1.Cls
        Label20.Visible = True
        lastButtonPressed = 1

       dateOfBirth = Format(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text, "dd/MM/yyyy")
       dateOfCondition = Format(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text, "dd/MM/yyyy")

        'FINDS THE AGE OF PERSON IN DAYS ON CONDITION DATE
        
        ndays = DateDiff("d", dateOfBirth, dateOfCondition)
        
        'FINDS THE EXACT POSITION OF EACH CYCLE
        l = (ndays / physicalCycleDays - Int(ndays / physicalCycleDays)) * physicalCycleDays
        m = (ndays / emotionalCycleDays - Int(ndays / emotionalCycleDays)) * emotionalCycleDays
        n = (ndays / intellectualCycleDays - Int(ndays / intellectualCycleDays)) * intellectualCycleDays
         
        b = (16 - l) * days
        r = (16 - m) * days
        g = (16 - n) * days
        
        'DRAWS WAVES FROM CENTER TO FORWARD
        Call MakeCurve(days + b, TopY, (days * physicalCycleDays) + b, RGB(255, 0, 0))
        Call MakeCurve(days + r, TopY, (days * emotionalCycleDays) + r, RGB(0, 255, 0))
        Call MakeCurve(days + g, TopY, (days * intellectualCycleDays) + g, RGB(0, 0, 255))
        
        'DRAWS WAVES FROM CENTER TO BACKWARD
        Call MakeReverseCurve(days + b, TopY, (days * physicalCycleDays) + b, RGB(255, 0, 0))
        Call MakeReverseCurve(days + r, TopY, (days * emotionalCycleDays) + r, RGB(0, 255, 0))
        Call MakeReverseCurve(days + g, TopY, (days * intellectualCycleDays) + g, RGB(0, 0, 255))
Else    'COMPARE MODE

        'CHECKS FOR VALIDITY OF DATE
        If Not (IsDate(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text) And IsDate(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text) And IsDate(Combo12.Text & "/" & Combo13.Text & "/" & Combo14.Text)) Then
            MsgBox "The Entered date is not valid."
            Exit Sub
        End If
        
        Picture1.Cls
        Label20.Visible = True
        lastButtonPressed = 1
        dateOfBirthCompare = Format(Combo12.Text & "/" & Combo13.Text & "/" & Combo14.Text, "dd/mm/yyyy")
        dateOfBirth = Format(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text, "dd/mm/yyyy")
        dateOfCondition = Format(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text, "dd/mm/yyyy")
        
        
        ndays = DateDiff("d", dateOfBirth, dateOfCondition)
        
        'FINDS THE EXACT POSITION OF EACH CYCLE
        l = (ndays / physicalCycleDays - Int(ndays / physicalCycleDays)) * physicalCycleDays
        m = (ndays / emotionalCycleDays - Int(ndays / emotionalCycleDays)) * emotionalCycleDays
        n = (ndays / intellectualCycleDays - Int(ndays / intellectualCycleDays)) * intellectualCycleDays
         
        b = (16 - l) * days
        r = (16 - m) * days
        g = (16 - n) * days
        
        
        ndaysCompare = DateDiff("d", dateOfBirthCompare, dateOfCondition)
        
        'FINDS THE EXACT POSITION OF EACH CYCLE
        lc = (ndaysCompare / physicalCycleDays - Int(ndaysCompare / physicalCycleDays)) * physicalCycleDays
        mc = (ndaysCompare / emotionalCycleDays - Int(ndaysCompare / emotionalCycleDays)) * emotionalCycleDays
        nc = (ndaysCompare / intellectualCycleDays - Int(ndaysCompare / intellectualCycleDays)) * intellectualCycleDays
         
        bc = (16 - lc) * days
        rc = (16 - mc) * days
        gc = (16 - nc) * days
        
        If dateOfBirthCompare > dateOfBirth Then
        
        Label3(2).Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, physicalCycleDays), 0) & "%"
        TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, physicalCycleDays)) + TopY
        Call MakeCurveCompare(days + bc, TopYCompare, ((days * physicalCycleDays) / 2) + bc, RGB(255, 0, 0))
        Call MakeReverseCurveCompare(days + bc, TopYCompare, ((days * physicalCycleDays) / 2) + bc, RGB(255, 0, 0))
        
        Label3(4).Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, emotionalCycleDays), 0) & "%"
        TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, emotionalCycleDays)) + TopY
        Call MakeCurveCompare(days + rc, TopYCompare, ((days * emotionalCycleDays) / 2) + rc, RGB(0, 255, 0))
        Call MakeReverseCurveCompare(days + rc, TopYCompare, ((days * emotionalCycleDays) / 2) + rc, RGB(0, 255, 0))
        
        Label3(0).Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intellectualCycleDays), 0) & "%"
        TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intellectualCycleDays)) + TopY
        Call MakeCurveCompare(days + gc, TopYCompare, ((days * intellectualCycleDays) / 2) + gc, RGB(0, 0, 255))
        Call MakeReverseCurveCompare(days + gc, TopYCompare, ((days * intellectualCycleDays) / 2) + gc, RGB(0, 0, 255))
        
        Else
        Label3(2).Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, physicalCycleDays), 0) & "%"
        
        TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, physicalCycleDays)) + TopY
        Call MakeCurveCompare(days + b, TopYCompare, ((days * physicalCycleDays) / 2) + b, RGB(255, 0, 0))
        Call MakeReverseCurveCompare(days + b, TopYCompare, ((days * physicalCycleDays) / 2) + b, RGB(255, 0, 0))
        Label3(4).Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, emotionalCycleDays), 0) & "%"

        TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, emotionalCycleDays)) + TopY
        Call MakeCurveCompare(days + r, TopYCompare, ((days * emotionalCycleDays) / 2) + r, RGB(0, 255, 0))
        Call MakeReverseCurveCompare(days + r, TopYCompare, ((days * emotionalCycleDays) / 2) + r, RGB(0, 255, 0))
        Label3(0).Caption = Round(GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intellectualCycleDays), 0) & "%"

        TopYCompare = (-10.8 * GetPeakDiff(dateOfBirth, dateOfBirthCompare, dateOfCondition, intellectualCycleDays)) + TopY
        Call MakeCurveCompare(days + g, TopYCompare, ((days * intellectualCycleDays) / 2) + g, RGB(0, 0, 255))
        Call MakeReverseCurveCompare(days + g, TopYCompare, ((days * intellectualCycleDays) / 2) + g, RGB(0, 0, 255))
 
        End If
       
       
End If

Label20.Visible = True
Label39.Visible = True
Label21.Visible = True

End Select
End Sub

Private Sub Label40_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call MoveEffect(Label40, Index, 11)
End Sub



Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'SHOWS DATE AND TIME WHEN MOUSE MOVES OVER THE GRAPH
Label20.Caption = (dateOfCondition - (16 - (x / days)))
If y >= 120 And y <= 4440 Then
Label39.Caption = PercentLang & ": " & Round(((2280 - y) / 2160) * 100, 0) & "%"
End If
'THE MOVE EFFECT
Call MoveEffect(Label40, 32, 11)
End Sub


'POSITIONS THE FORM TO THE CENTER
Private Sub OptimizeFormPosition()
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = ((Screen.Height - 322) / 2) - (Me.Height / 2)

End Sub
Public Sub List_Load(thelist As ListBox, FileName As String)
'LOADS A FILE TO A LIST BOX
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        If TheContents$ = "" Then
        Else
        Call List_Add(thelist, TheContents$)
        End If
    Loop Until EOF(fFile)
    Close fFile
End Sub

Public Sub List_Save(thelist As ListBox, FileName As String)
    'SAVE A LISTBOX AS FILENAME
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Output As fFile
    For Save = 0 To thelist.ListCount - 1
        Print #fFile, List1.list(Save)
    Next Save
    Close fFile
End Sub
Public Sub List_Add(list As ListBox, txt As String)
On Error Resume Next
    List1.AddItem txt
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("|") Then
KeyAscii = 0
MsgBox "The charachter '|' is a sysntax indentifier."
End If
End Sub

Private Sub ShowGuideMeCallOut(Message As String, PositionOfObj As Object)
        If Me.Width = 11820 Then
        Form1.Label1 = Message
        Form1.Top = Me.Top + PositionOfObj.Top + 100 + Frame1.Top
        Form1.Left = Me.Left + PositionOfObj.Left + Frame1.Left
        Load Form1
        Form1.Show
        Else
        Form1.Label1 = Message
        Form1.Top = Me.Top + PositionOfObj.Top + 100
        Form1.Left = Me.Left + PositionOfObj.Left
        Load Form1
        Form1.Show
        End If
End Sub
Sub MoveEffect(obj As Object, Index As Integer, Ubnd As Integer)
Dim i As Single
If moveconstant <> Index Then
For i = 0 To Ubnd
If i = Index Then
obj(i).FontBold = True
obj(i).BorderStyle = 1
LastMoveEffect = i
Exit For
Else
obj(LastMoveEffect).FontBold = False
obj(LastMoveEffect).BorderStyle = 0
End If
Next
End If
moveconstant = Index
End Sub
Private Sub SetLanguage(lang As String)
Dim i As Integer
Select Case (lang)
Case ""
Me.Caption = "Techno Rex BioRhythm Version 3.5"
S1 = "SELECT NAME of person OR Enter Date Of Birth of Person. "
S2 = "CLICK here to put Todays Date in Condition Combo Box."
S3 = "CLICK here to Add a perons name."
S4 = "CLICK here to Remove a perons name."
S5 = "SELECT NAME of person OR Enter Date Of Birth of Person. "
S6 = "CLICK Any of these three buttons."
S7 = "SELECT NAME of person OR Enter Date Of Birth of Person. "
S8 = "CLICK HERE to Switch OFF Compare Mode."
S9 = "CLICK Any of these three buttons."
S10 = "CLICK here to Switch ON compare mode"
S11 = "SELECT NAME of person OR Enter Date Of Birth of Person. "
S12 = "CLICK any of these three buttons."
S13 = "Write name of person and enter the date of birth and then CLICK HERE "
S14 = "Select name of the person from list and then CLICK here."

m1 = "Please Select a name to Remove"
m2 = "Entered date is not valid"
OffButton = "OFF"
OnButton = "ON"
PercentLang = "Percentage"
UBButton = "Understanding Biorhythm "
ARButton = "Add/Remove Name      "
HelpButton = "               Help               "

WBiorhythm = "What is Biorhythm?"
ABiorhythm = "Biorhythm study and use is considered a ""pseudo science"" in the United State however it is widely accepted and utilized throughout Europe and much of the rest of the world.              " & _
             "Biorhythms are inherent cycles which regulate your metabolism, coordination, emotions, memory, and more.  As your biorhythm cycles rise and fall, so does your ability to perform physical activities, deal with stress, and make sound decisions."
WPhyCycle = "Meaning of Physical Cycle"
APhyCycle = "The physical cycle is 23 days long and is the dominant cycle in men. It regulates hand-eye coordination, strength, endurance, sex drive, initiative, metabolic rate, resistance to, and recovery from illness. Surgery should be avoided on physical transition days and during negative physical cycles. "

WEmoCycle = "Meaning of Emotional Cycle"
AEmoCycle = "The emotional cycle is 28 days long andis the dominant cycle in women. It regulates emotions, feelings, mood, sensitivity, sensation, sexuality, fantasy, temperament, nerves, reactions, affections and creativity."

WIntCycle = "Meaning of Intellectual Cycle"
AIntCycle = "The intellectual cycle is 33 days long and regulates intelligence, logic, mental reaction, alertness, sense of direction, decision-making, judgment, power of deduction, memory, and ambition."

WPassCycle = "Meaning of Passion Cycle"
APassCycle = "Passion cycle is the composite of the Physical and Emotional cycles. Passion encompasses your motivation to act, and the drive that allows you to continue a difficult pursuit. This cycle also tracks sexuality in its purest form."

WMastCycle = "Meaning of Mastery Cycle"
AMastCycle = "Mastery Cycle is the composite of the Intellectual and Physical cycles. Mastery encompasses your ability to succeed at tasks and to obtain what you desire. This cycle also tracks athletic ability and the focus required to learn physical skills."

WWisdCycle = "Meaning of Wisdom Cycle"
AWisdCycle = "Wisdom Cycle is  the composite of the Emotional and Intellectual cycles. Wisdom encompasses your understanding of the world, your role in it, and the things that are truly important to your life. This cycle also tracks the presence of mind that you need to make crucial decisions."

WIntuCycle = "Meaning of Intutive Cycle"
AIntuCycle = "Intuitive Cycle is of 38 days and shows your intuition or sixth sense."

WAesCycle = "Meaning of Aesthetic Cycle"
AAesCycle = "Aesthetic cycle is 43 days long and describes interest in the beautiful and the harmonious."

WSelfCycle = "Meaning of Self Awareness Cycle"
ASelfCycle = " Self-Awareness - 48 days, it expresses ability to percept own personality and individuality."

WSpirCycle = "Meaning of Spiritual Cycle"
ASpirCycle = "Spiritual Cycle is 53 days long and  describes inner stability and relaxed attitude. "

WNP = "Negative and Positive 100%"
ANP = "The numbers from +100% (maximum) to -100% (minimum) indicate where the rhythms are on a particular day. In general, a rhythm at 0% is thought to have no real impact on your life, whereas a rhythm at +100% (a high) would give you an edge in that area, and a rhythm at -100% (a low) would make life more difficult in that area. There is no particular meaning to a day on which your rhythms are all high or all low, except the obvious benefits or hindrances that these rare extremes are thought to have on your life."

WBal = "Balance is the key"
ABal = "Understanding your positive cycles may assist you in planning surgery, physical outings, sporting events, exams, and job interviews. " & _
            "Understanding your negative cycles and your particular reaction to them…may help you avoid accidents, hurtful situations, unnecessary grief and misfortune."
Case Else
Open App.Path & "\Language\" & lang & ".lng" For Input As #1
i = 0
Line Input #1, T(88)
Line Input #1, T(86)
Line Input #1, T(87)
Do While Not EOF(1)
Line Input #1, T(i)
i = i + 1
Loop

Me.Caption = T(1)
PercentLang = T(2)
S1 = T(3)
S2 = T(4)
S3 = T(5)
S4 = T(6)
S5 = T(7)
S6 = T(8)
S7 = T(9)
S8 = T(10)
S9 = T(11)
S10 = T(12)
S11 = T(13)
S12 = T(14)
S13 = T(15)
S14 = T(16)
OffButton = T(17)
OnButton = T(18)
UBButton = T(19)
ARButton = T(20)
HelpButton = T(21)

Label41.Caption = T(22)
Label30.Caption = T(23)
Label34.Caption = T(23)
Label25.Caption = T(23)


Label2.Caption = T(24)
Label33.Caption = T(24)
Label26.Caption = T(24)

Label1.Caption = T(25)

Label11.Caption = T(26)
Label27.Caption = T(26)
Label12.Caption = T(27)
Label28.Caption = T(27)
Label13.Caption = T(28)
Label29.Caption = T(28)

Label23.Caption = T(29)

Label40(7).Caption = T(30)

Label40(11).Caption = T(31)
Label40(10).Caption = T(32)
Label40(9).Caption = T(33)


Label40(3).Caption = T(34)
Label37.Caption = T(35)


Label40(2).Caption = T(36)
Label40(1).Caption = T(37)

Label35.Caption = T(38)

Label32.Caption = T(39)
Label36.Caption = T(40)


Label3(1).Caption = T(41)
Label3(3).Caption = T(42)
Label3(5).Caption = T(43)
Label5.Caption = T(44)
Label7.Caption = T(45)
Label9.Caption = T(46)
Label15.Caption = T(47)
Label22.Caption = T(48)
Label17.Caption = T(49)
Label19.Caption = T(50)
Label21.Caption = T(51)


WBiorhythm = T(52)
ABiorhythm = T(53)
WPhyCycle = T(54)
APhyCycle = T(55)

WEmoCycle = T(56)
AEmoCycle = T(57)

WIntCycle = T(58)
AIntCycle = T(59)

WPassCycle = T(60)
APassCycle = T(61)

WMastCycle = T(62)
AMastCycle = T(63)

WWisdCycle = T(64)
AWisdCycle = T(65)

WIntuCycle = T(66)
AIntuCycle = T(67)

WAesCycle = T(68)
AAesCycle = T(69)

WSelfCycle = T(70)
ASelfCycle = T(71)

WSpirCycle = T(72)
ASpirCycle = T(73)

WNP = T(74)
ANP = T(75)

WBal = T(76)
ABal = T(77)


Label38(0).Caption = T(78)
Label38(1).Caption = T(79)
Label38(2).Caption = T(80)
Label38(6).Caption = T(81)
Label38(7).Caption = T(82)
Label40(8).Caption = T(83)
m1 = T(84)
m2 = T(85)

Call ConvertFont(Me, 0, T(86), T(87))
Call ConvertFont(Label1, 0, T(86), T(87))
Call ConvertFont(Label2, 0, T(86), T(87))
Call ConvertFont(Label3, Label3.UBound, T(86), T(87))
Call ConvertFont(Label4, 0, T(86), T(87))
Call ConvertFont(Label5, 0, T(86), T(87))
Call ConvertFont(Label6, 0, T(86), T(87))
Call ConvertFont(Label7, 0, T(86), T(87))
Call ConvertFont(Label8, 0, T(86), T(87))
Call ConvertFont(Label9, 0, T(86), T(87))
Call ConvertFont(Label10, 0, T(86), T(87))
Call ConvertFont(Label11, 0, T(86), T(87))
Call ConvertFont(Label12, 0, T(86), T(87))
Call ConvertFont(Label13, 0, T(86), T(87))
Call ConvertFont(Label14, 0, T(86), T(87))
Call ConvertFont(Label15, 0, T(86), T(87))
Call ConvertFont(Label16, 0, T(86), T(87))
Call ConvertFont(Label17, 0, T(86), T(87))
Call ConvertFont(Label18, 0, T(86), T(87))
Call ConvertFont(Label19, 0, T(86), T(87))
Call ConvertFont(Label20, 0, T(86), T(87))
Call ConvertFont(Label21, 0, T(86), T(87))
Call ConvertFont(Label22, 0, T(86), T(87))
Call ConvertFont(Label23, 0, T(86), T(87))
Call ConvertFont(Label24, Label24.UBound, T(86), T(87))
Call ConvertFont(Label25, 0, T(86), T(87))
Call ConvertFont(Label26, 0, T(86), T(87))
Call ConvertFont(Label27, 0, T(86), T(87))
Call ConvertFont(Label28, 0, T(86), T(87))
Call ConvertFont(Label29, 0, T(86), T(87))
Call ConvertFont(Label30, 0, T(86), T(87))
Call ConvertFont(Label32, 0, T(86), T(87))
Call ConvertFont(Label33, 0, T(86), T(87))
Call ConvertFont(Label34, 0, T(86), T(87))
Call ConvertFont(Label35, 0, T(86), T(87))
Call ConvertFont(Label36, 0, T(86), T(87))
Call ConvertFont(Label37, 0, T(86), T(87))
Call ConvertFont(Label38, Label38.UBound, T(86), T(87))
Call ConvertFont(Label39, 0, T(86), T(87))
Call ConvertFont(Label40, Label40.UBound, T(86), T(87))
Call ConvertFont(Label41, 0, T(86), T(87))
If language = "Hindi" Then
Call ConvertFont(Combo1, 0, T(86), 12)
Call ConvertFont(Combo2, 0, T(86), 12)
Call ConvertFont(Combo3, 0, T(86), 12)
Call ConvertFont(Combo4, 0, T(86), 12)
Call ConvertFont(Combo5, 0, T(86), 12)
Call ConvertFont(Combo6, 0, T(86), 12)
Call ConvertFont(Combo8, 0, T(86), 12)
Call ConvertFont(Combo9, 0, T(86), 12)
Call ConvertFont(Combo10, 0, T(86), 12)
Call ConvertFont(Combo12, 0, T(86), 12)
Call ConvertFont(Combo13, 0, T(86), 12)
Call ConvertFont(Combo14, 0, T(86), 12)
End If
Label40(6).Caption = UBButton
Label40(5).Caption = ARButton
Label40(4).Caption = HelpButton
Label40(0).Caption = OffButton
End Select
End Sub
Private Sub ConvertFont(obj As Object, Ubnd As Integer, fontname As String, fontsize As String)
On Error Resume Next
Dim i As Integer
If Ubnd = 0 Then
obj.fontname = fontname
obj.fontsize = Val(fontsize)
Else
For i = 0 To Ubnd
obj(i).fontname = fontname
obj(i).fontsize = Val(fontsize)
Next
End If
End Sub
Public Function GetPeakDiff(d1 As Date, D2 As Date, ConD As Date, Cycle As Integer) As Double

 
 Dim diff1 As Double
 Dim diff2 As Double
 Dim r1 As Double
 Dim r2 As Double

 
 diff1 = ((DateDiff("d", d1, ConD)) / Cycle)
 diff2 = ((DateDiff("d", D2, ConD)) / Cycle)
 
 r1 = (Abs(diff1 - diff2) - Int(Abs(diff1 - diff2))) * Cycle
 
 GetPeakDiff = Cos(r1 * (3.14 / Cycle)) * 100
 End Function

'MAKE CURVES FORWARDS COMPARE
 Public Sub MakeCurveCompare(x As Double, y As Double, d As Double, c As Long)
 Dim T
 Picture1.ForeColor = c
 Picture1.PSet (x, y)
 For T = x To days * 31
    Picture1.Line -(T, y - (1080 * Sin((3.14) * ((T - x) / ((d - x) / 2)))))
 Next
End Sub
'MAKES CURVES BACKWARDS COMPARE
Public Sub MakeReverseCurveCompare(x As Double, y As Double, d As Double, c As Long)
Dim T
Picture1.ForeColor = c
Picture1.PSet (days - d + x, y)
For T = (days - d + x) To x
    Picture1.Line -(T, y - (1080 * Sin((3.14) * ((T - x) / ((d - x) / 2)))))
Next
End Sub

'MAKES SECONDARY CURVE COMPARE
Public Sub MakeSecCurveCompare(x1 As Double, x2 As Double, x1c As Double, x2c As Double, y1 As Double, y2 As Double, y1c As Double, y2c As Double, D2 As Double, d1 As Double, D2c As Double, D1c As Double, c1 As Long)
Dim T, y
 Picture1.ForeColor = c1
 y = ((y1 - (1080 * Sin((3.14) * ((T - x1) / ((d1 - x1) / 2))))) + (y2 - (1080 * Sin((3.14) * ((T - x2) / ((D2 - x2) / 2)))))) + ((y1c - (1080 * Sin((3.14) * ((T - x1c) / ((D1c - x1c) / 2))))) + (y2c - (1080 * Sin((3.14) * ((T - x2c) / ((D2c - x2c) / 2))))))
 Picture1.PSet (0, y / 4)
 For T = 0 To days * 31
 y = ((y1 - (1080 * Sin((3.14) * ((T - x1) / ((d1 - x1) / 2))))) + (y2 - (1080 * Sin((3.14) * ((T - x2) / ((D2 - x2) / 2)))))) + ((y1c - (1080 * Sin((3.14) * ((T - x1c) / ((D1c - x1c) / 2))))) + (y2c - (1080 * Sin((3.14) * ((T - x2c) / ((D2c - x2c) / 2))))))
 Picture1.Line -(T, y / 4)
 Next
End Sub
Private Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        Open FullFileName For Input As #1
        Close #1
        FileExists = True
    Exit Function
MakeF:
        FileExists = False
    Exit Function
End Function

