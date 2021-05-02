VERSION 5.00
Begin VB.Form detailForm 
   Caption         =   "Detail Information"
   ClientHeight    =   11265
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   14700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   751
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_More 
      Caption         =   "more"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   10560
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   5880
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   9
      Top             =   720
      Width           =   3870
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3615
      Left            =   14160
      Max             =   584
      TabIndex        =   8
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   9960
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   7
      Top             =   720
      Width           =   4245
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   19965
         Left            =   0
         ScaleHeight     =   1329
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   281
         TabIndex        =   10
         Top             =   0
         Width           =   4245
      End
   End
   Begin VB.CommandButton cmdReduction 
      Caption         =   "default"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   5175
      Begin VB.Image Image1 
         Height          =   3615
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdExpand 
      Caption         =   "Expand"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   975
   End
   Begin VB.PictureBox graPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   240
      ScaleHeight     =   335
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   943
      TabIndex        =   1
      Top             =   5400
      Width           =   14175
      Begin VB.Label lbl_Class 
         Caption         =   "Class"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl_Length 
         Caption         =   "Template Length"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lbl_id 
         Caption         =   "Id"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "close"
      Height          =   375
      Left            =   13200
      TabIndex        =   0
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   72
      X2              =   952
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label Label2 
      Caption         =   "Iris Data"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   960
      Y1              =   344
      Y2              =   344
   End
   Begin VB.Label Label1 
      Caption         =   "Pattern Distribution"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   5040
      Width           =   2055
   End
End
Attribute VB_Name = "detailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ILeft As Long
Dim ITop As Long
Dim IWidth As Long
Dim IHeight As Long
Dim defaultWidth As Long
Dim defaultHeight As Long
Dim PicWidth As Long
Dim PicHeight As Long
Dim isExpanded As Boolean
Dim mX, mY As Single
Dim isDraging As Boolean
Dim TotMoveX As Single
Dim TotMoveY As Single



Private Sub cmd_More_Click()
  
    frmFull.img_Picture.Picture = Image1.Picture
   
    
    
    frmFull.Show 1
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdExpand_Click()
Image1.Left = ILeft - 100
Image1.Top = ITop - (100 * 0.75)
Image1.Width = Image1.Width + 200
Image1.Height = Image1.Height + (200 * 0.75)

ILeft = Image1.Left
ITop = Image1.Top

IWidth = Image1.Width
IHeight = Image1.Height

isExpanded = True
    
    
End Sub



Private Sub Form_Load()
    ILeft = Image1.Left
    ITop = Image1.Top
    IWidth = Image1.Width
    IHeight = Image1.Height
    Picture1.Picture = LoadPicture("1.bmp")
    Picture2.Picture = LoadPicture("2.bmp")
    defaultWidth = Image1.Width
    defaultHeight = Image1.Height
    graPic.BackColor = RGB(50, 0, 0)
    PicWidth = graPic.Width
    PicHeight = graPic.Height
    isExpanded = False
    isDraging = False
    TotMoveX = 0
    TotMoveY = 0
    
    InitGraph
    RenderToHeaderAndData
   
End Sub

Public Sub RenderToHeaderAndData()
    Dim cnt, i As Long
    Dim Name, Info As String
    Dim fBin, bTmp As Byte
    Dim fileName As String
    Dim INFO_TOT_SIZE As Long '4Byte NETotalSize
    Dim INFO_TOT_CLASS As Byte '1byte NETotalClass
    
    Dim IrisData() As Byte
    Dim offset As Long
    Dim k As Long
    
    
    fileName = "REG_" & selectId & ".bin" ' 전달된 아이디 값으로 파일명 생성
    lbl_id.Visible = True
    lbl_Length.Visible = True
    lbl_Class.Visible = True
    
    If (Dir(fileName) = "") Then
        
    End If
    fBin = FreeFile()
    Open fileName For Binary Access Read As fBin
    'Header Informatin read---------------------------------------'
     For cnt = 0 To 19
        Get fBin, , bTmp
        Info = Info + Chr(bTmp)
        Next cnt
        
        'Name 20 byte
        For cnt = 0 To 19
            Get fBin, , bTmp
            Name = Name + Chr(bTmp)
        Next cnt
        
        lbl_id.Caption = "ID : " & Name
        
        Get fBin, , INFO_TOT_SIZE
        Get fBin, , INFO_TOT_CLASS
        
        ReDim IrisData(0 To INFO_TOT_SIZE - 1) '템플릿 길이에 따라 동적 배열 생성
        
        lbl_Length.Caption = "Tmeplate Length : " & Str(INFO_TOT_SIZE)
        lbl_Class.Caption = "Class : " & Str(INFO_TOT_CLASS)
        lbl_id.Refresh
        lbl_Length.Refresh
        lbl_Class.Refresh
        Get fBin, , IrisData
    Close fBin
  
    offset = INFO_TOT_SIZE / 900
    
    For k = 0 To 900
        If k > 900 Then
            Exit Sub
        End If
        graPic.Line (20 + k + 2, PicHeight - 20 - IrisData(k))-(20 + k + 2 + 1, PicHeight - 20 - IrisData(k + 1)), RGB(255, 255, 153)
    Next k
    
    
    
End Sub
Public Sub InitGraph()

    lbl_id.BackColor = RGB(50, 0, 0)
    lbl_Length.BackColor = RGB(50, 0, 0)
    lbl_Class.BackColor = RGB(50, 0, 0)
    
    
    graPic.Line (20, 35)-(PicWidth - 20, PicHeight - 20), QBColor(0), BF
    graPic.DrawWidth = 1
    graPic.Line (20, 35)-(PicWidth - 20, PicHeight - 20), RGB(131, 119, 108), B
End Sub

Public Sub InitGraph2()
    
    graPic.DrawWidth = 1
    graPic.Line (20, 35)-(PicWidth - 20, PicHeight - 20), RGB(131, 119, 108), B
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    
    mX = X
    mY = Y
    isDraging = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lx, ly As Single
    DoEvents
    If isDraging = False Then Exit Sub
    If (Image1.Width < defaultWidth + 1400) Then Exit Sub
    
    lx = mX - X
    ly = mY - Y
    TotMoveX = TotMoveX + lx
    TotMoveY = TotMoveY - ly
    
    
    If (Image1.Left - lx > 0 Or Image1.Top - ly > 0) Then Exit Sub
    If ((Image1.Left + Image1.Width - lx < 5715) Or (Image1.Top + Image1.Height - ly < 3615)) Then Exit Sub
    
    Image1.Left = Image1.Left - lx
    Image1.Top = Image1.Top - ly
    
    
   
    
    
    mX = X
    mY = Y
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isDraging Then
        isDraging = False
    End If
End Sub

Private Sub VScroll1_Change()
    Picture2.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    Picture2.Top = -VScroll1.Value
End Sub
