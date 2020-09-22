VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   735
      Left            =   120
      Picture         =   "frmCallOutHelp.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "hoeto fet to dhausjdf coddfdfdfdfkdfsdfsdkfdskfsdkfdk xxxxxxxxxxxxxx"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   105
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Picture         =   "frmCallOutHelp.frx":4942
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    bottom As Long
End Type


Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2




Dim cnt As Byte
Private Function GetTextRgn() As Long
    Dim hRgn1 As Long, hRgn2 As Long
    Dim rct As RECT
    
    BeginPath hdc

    TextOut hdc, 10, 10, Chr$(120), 1     'Windows Flag

    'Circle (2000, 2000), 1000            'Circle window

'Create any path you want in this section to create your irregular window.
  
    Form1.Line (100, 1075)-(100 + Label1.Width, 1075 + Label1.Height), , BF
 
    EndPath hdc
    
    hRgn1 = PathToRegion(hdc)
    GetRgnBox hRgn1, rct
    hRgn2 = CreateRectRgnIndirect(rct)
    CombineRgn hRgn2, hRgn2, hRgn1, 1
    DeleteObject hRgn1
    GetTextRgn = hRgn2
End Function



Private Sub Form_Load()
    Dim hRgn As Long
    cnt = 0
    Call putMeOnTop(Me)

    Me.Font.Name = "Wingdings 3"
    Me.Font.Size = 50
    
    hRgn = GetTextRgn()
    SetWindowRgn hwnd, hRgn, 1
 
End Sub

Private Sub Form_LostFocus()
 Unload Me
End Sub

Public Sub putMeOnTop(Form As Form)
    SetWindowPos Form.hwnd, -1, 0, 0, 0, 0, 1 Or 2
End Sub
