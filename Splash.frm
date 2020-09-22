VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4950
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400040&
      Height          =   2175
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "English"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   120
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   120
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   120
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   120
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jai Shri Ram"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Programmer Anshuk Kumar                                     Contact:   Anshukk@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   4200
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   240
      Picture         =   "Splash.frx":1CCA
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   4980
      Left            =   0
      Picture         =   "Splash.frx":442F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7380
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moveconstant As Single
Dim LastMoveEffect As Single
Dim language() As String
Dim Declaration() As String
Dim FName() As String
Dim FSize() As String
Dim FileName() As String

Private Sub Form_Load()
On Error GoTo err:
Dim i As Single
Dim nfiles As Single
restart:
moveconstant = 32
File1.Pattern = "*.lng"
File1.Path = App.Path & "\Language"
If File1.ListCount >= 4 Then
nfiles = 4
Else
nfiles = File1.ListCount
End If
ReDim FileName(0 To nfiles) As String
ReDim language(0 To nfiles) As String
ReDim Declaration(0 To nfiles) As String
ReDim FSize(0 To nfiles) As String
ReDim FName(0 To nfiles) As String
language(0) = "English"
FName(0) = "ms sans serif"
FSize(0) = "8"
For i = 1 To nfiles
File1.ListIndex = i - 1
Open File1.Path & "\" & File1.FileName For Input As 1
If LOF(1) = 0 Then
Close #1
Kill (File1.Path & "\" & File1.FileName)
File1.Refresh
Shape1(i).Visible = False
Label6(i).Visible = False
GoTo restart
End If
FileName(i) = Mid(File1.FileName, 1, Len(File1.FileName) - 4)
Line Input #1, language(i)
Line Input #1, FName(i)
Line Input #1, FSize(i)
Line Input #1, Declaration(i)
Close #1
Label6(i).fontname = FName(i)
Label6(i).fontsize = FSize(i)
Label6(i).Caption = language(i)
Shape1(i).Visible = True
Label6(i).Visible = True
Next
DisplayLanguage (0)
Exit Sub
err:
MsgBox "Invalid Language file : " & FileName(i) & ".lng"
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call MoveEffect(Label6, 32, UBound(language))
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call MoveEffect(Label6, 32, UBound(language))
End Sub

Private Sub Label6_Click(Index As Integer)
MainFrm.language = FileName(Index)
Load MainFrm
MainFrm.Show
Unload Me
End Sub
Sub MoveEffect(obj As Object, Index As Integer, Ubnd As Integer)
Dim i As Single
If moveconstant <> Index Then
For i = 0 To Ubnd
If i = Index Then
obj(i).FontBold = True
obj(i).BorderStyle = 1
DisplayLanguage (Index)
Else
obj(i).FontBold = False
obj(i).BorderStyle = 0

End If
Next
End If
moveconstant = Index
End Sub
Private Sub Label6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call MoveEffect(Label6, Index, UBound(language))
End Sub
Private Sub DisplayLanguage(lang As Integer)
Select Case (lang)
Case 0
Label4.Caption = "This software may only be used for educational and Informative purpose and can  be distributed, changed or translated. However if you use it for  commercial purpose please pay $25.00 (US). For details contact   Anshukk@yahoo.com  .I take no responsiblity for any damage that may occur as a result of the use of the included source code and/or the other support files associated with it."
Case Else
Label4.Caption = Declaration(lang)
End Select
Label4.fontname = FName(lang)
Label4.fontsize = FSize(lang)
End Sub
