VERSION 5.00
Begin VB.Form frmFull 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   13575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16890
   FillColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   905
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1126
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox img_Picture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   600
      ScaleHeight     =   3585
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   1560
      Width           =   4800
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "INFO"
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   5160
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   3600
         X2              =   3600
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   1560
         X2              =   1560
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Label Label3 
         Caption         =   "CLASS:"
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "TEMPLATE:"
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "ID :"
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox tmpPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   480
      ScaleHeight     =   318
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   2
      Top             =   9600
      Visible         =   0   'False
      Width           =   4815
      Begin VB.PictureBox sol_Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4800
         Left            =   0
         ScaleHeight     =   318
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   318
         TabIndex        =   3
         Top             =   0
         Width           =   4800
      End
   End
   Begin VB.PictureBox sol_Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   480
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   384
      X2              =   384
      Y1              =   40
      Y2              =   136
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   3915
      Left            =   360
      Top             =   1320
      Width           =   5205
   End
End
Attribute VB_Name = "frmFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P_SCREEN_WIDTH As Long  '픽셀로 변환된 화면 스크린의 폭
Dim P_SCREEN_HEIGHT As Long '픽셀로 변환된 스크린의 높이
Dim g_Width As Long         '그래프가 그려질 공간의 폭
Dim g_Height As Long        '그래프가 그려질 공간의 높이( 1개의 높이임)
Dim g_Total_Height As Long  '그래프가 그려질 전체 공간의 높이
Dim g_rate As Single        '실제 스크린 크기에따라 축소 비율
Dim PIXEL_TO_RATE As Long
Dim offset As Long          '화면상에서 그려진 그래프 의 갯수 ( 데이타 길이 / 그래프 공간의 폭)
Dim IrisData() As Byte      '템플릿 데이타가 저장되는 배열,
Dim INFO_TOT_SIZE As Long '4Byte NETotalSize
Dim INFO_TOT_CLASS As Byte '1byte NETotalClass


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
   
    Width = Screen.Width '폼의 폭을 화면의 폭으로 맞춤
    Height = Screen.Height '폼의 높이를 화면의 높이로 맞춤
   
    Top = 0 '폼의 좌표를 화면 상단 왼쪽으로 맞춤
    Left = 0 '
    
    sol_Picture1.Picture = LoadPicture("1_1.bmp") ' 이미지 로드
    sol_Picture2.Picture = LoadPicture("2_2.bmp") ' 이미지 로드
    
    P_SCREEN_WIDTH = Screen.Width / 15 '스크린의 트윕폭을 픽셀로 변환
    P_SCREEN_HEIGHT = Screen.Height / 15 '스크린의 트윕 높이을 픽셀로 변환
    
     
    
    g_Height = 270 ' 그래프가 그려질 작업공간의 height 설정 - 그래프 1개의 높이
    g_Width = P_SCREEN_WIDTH - 400 - 20 ' 그래프가 그려질 작업공간의 width 설정
    
    
    
    Line1.BorderColor = RGB(131, 119, 108)
    Line1.Y2 = P_SCREEN_HEIGHT - 40
    
    
    img_Picture.Top = 40 + 49 + 10
    sol_Picture1.Top = img_Picture.Top + img_Picture.Height + 10
    tmpPic.Top = sol_Picture1.Top + sol_Picture1.Height + 10
    tmpPic.Height = 400
    g_Total_Height = P_SCREEN_HEIGHT - 80
    
    Initcontrol ' 컨트롤 초기화
    InitData
    GetPixeltoRate
    InitGraph (offset)
   
    DrawBound offset
    

End Sub
Public Sub Initcontrol()
    
    BackColor = RGB(50, 0, 0)
    Frame1.BackColor = RGB(50, 0, 0)
    'Frame1.ForeColor = RGB(100, 0, 0)
    Label1.BackColor = RGB(50, 0, 0)
    Label2.BackColor = RGB(50, 0, 0)
    Label3.BackColor = RGB(50, 0, 0)
    
    Label1.ForeColor = RGB(255, 255, 255)
    Label2.ForeColor = RGB(255, 255, 255)
    Label3.ForeColor = RGB(255, 255, 255)
End Sub
Public Sub InitData()
    Dim cnt, i As Long
    Dim Name, Info As String
    Dim fBin, bTmp As Byte
    Dim fileName As String
    Dim k As Long
    fileName = "REG_" & selectId & ".bin" ' 전달된 아이디 값으로 파일명 생성
    
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
         
    Get fBin, , INFO_TOT_SIZE
    Get fBin, , INFO_TOT_CLASS
        
    Label1.Caption = "ID : " & Name
    Label2.Caption = "TEMPLATE : " & Str(INFO_TOT_SIZE)
    Label3.Caption = "CLASS : " & Str(INFO_TOT_CLASS)
    
    Label1.Refresh
    Label2.Refresh
    Label3.Refresh
        
    ReDim IrisData(0 To INFO_TOT_SIZE - 1) '템플릿 길이에 따라 동적 배열 생성
    Get fBin, , IrisData
    Close fBin
        
    offset = INFO_TOT_SIZE / g_Width
    If (INFO_TOT_SIZE Mod g_Width <> 0) Then
        
    End If
      
    
End Sub
Public Sub InitGraph(offset)
 
     Dim index As Long
    Dim i As Long
    Dim j As Long
    Dim v1 As Long
    Dim v2 As Long
    
    index = 0
    Dim r_v1 As Long
    Dim r_v2 As Long
    
    Dim startY
    startY = PIXEL_TO_RATE + 40
    
    For i = 0 To offset - 1

        Line (400, 40 + (i * PIXEL_TO_RATE + 5))-(P_SCREEN_WIDTH - 20, 40 + (i * PIXEL_TO_RATE) + PIXEL_TO_RATE), QBColor(0), BF
        DrawWidth = 1
        Line (400, 40 + (i * PIXEL_TO_RATE + 5))-(P_SCREEN_WIDTH - 20, 40 + (i * PIXEL_TO_RATE) + PIXEL_TO_RATE), RGB(131, 119, 108), B
     
         For j = 0 To g_Width - 1
            
            If (index >= INFO_TOT_SIZE - 3) Then
                Exit Sub
             End If
             
             v1 = IrisData(index)
             v2 = IrisData(index + 1)
             r_v1 = v1 * g_rate
             r_v2 = v2 * g_rate
             Line (400 + j, 40 + ((i * PIXEL_TO_RATE) + PIXEL_TO_RATE) - r_v1)-(400 + j + 1, 40 + ((i * PIXEL_TO_RATE) + PIXEL_TO_RATE) - r_v2), RGB(255, 255, 153)
            index = index + 1

         Next j
         startY = startY + (g_Height * g_rate)
    Next i
    
End Sub

Public Sub GetPixeltoRate()
    Dim i As Single
    For i = 1# To 0# Step -0.1
        If (i * 270 * offset < g_Total_Height) Then
            g_rate = i
            PIXEL_TO_RATE = i * g_Height
            Exit Sub
        End If
    Next i
    
End Sub


Public Sub DrawBound(ByVal offset As Long)
 Dim i As Long
    For i = 0 To offset - 1

        DrawWidth = 1
        Line (400, 40 + (i * PIXEL_TO_RATE) + 5)-(P_SCREEN_WIDTH - 20, 40 + (i * PIXEL_TO_RATE) + PIXEL_TO_RATE), RGB(131, 119, 108), B
    Next i

End Sub





