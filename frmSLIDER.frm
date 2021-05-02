VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSLIDER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IRIS FRAME LIST"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   20340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   670
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1356
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Pro 
      Height          =   300
      Left            =   7320
      TabIndex        =   1
      Top             =   810
      Visible         =   0   'False
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   945
      Left            =   240
      TabIndex        =   13
      Top             =   360
      Width           =   19935
      Begin VB.TextBox idBox 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   720
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblStatus 
         Caption         =   "Wait For Processing"
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "id"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSharpen 
      Caption         =   "Sharpen"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox impPictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3600
      Index           =   7
      Left            =   15360
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   10
      Top             =   5760
      Width           =   4800
      Begin VB.Line scanLine 
         BorderColor     =   &H0000FF00&
         Index           =   7
         Visible         =   0   'False
         X1              =   0
         X2              =   320
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox impPictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3600
      Index           =   6
      Left            =   10320
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   9
      Top             =   5760
      Width           =   4800
      Begin VB.Line scanLine 
         BorderColor     =   &H0000FF00&
         Index           =   6
         Visible         =   0   'False
         X1              =   0
         X2              =   320
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox impPictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3600
      Index           =   4
      Left            =   240
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   7
      Top             =   5760
      Width           =   4800
      Begin VB.Line scanLine 
         BorderColor     =   &H0000FF00&
         Index           =   4
         Visible         =   0   'False
         X1              =   0
         X2              =   320
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox impPictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3600
      Index           =   3
      Left            =   15360
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   6
      Top             =   1920
      Width           =   4800
      Begin VB.Line scanLine 
         BorderColor     =   &H0000FF00&
         Index           =   3
         Visible         =   0   'False
         X1              =   0
         X2              =   320
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox impPictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3600
      Index           =   2
      Left            =   10320
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   5
      Top             =   1920
      Width           =   4800
      Begin VB.Line scanLine 
         BorderColor     =   &H0000FF00&
         Index           =   2
         Visible         =   0   'False
         X1              =   0
         X2              =   320
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox impPictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3600
      Index           =   0
      Left            =   240
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   3
      Top             =   1920
      Width           =   4800
      Begin VB.Line scanLine 
         BorderColor     =   &H0000FF00&
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   320
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   840
      Top             =   9600
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "close"
      Height          =   375
      Left            =   18240
      TabIndex        =   0
      Top             =   9480
      Width           =   1935
   End
   Begin VB.PictureBox impPictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3600
      Index           =   1
      Left            =   5280
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   4
      Top             =   1920
      Width           =   4800
      Begin VB.Line scanLine 
         BorderColor     =   &H0000FF00&
         Index           =   1
         Visible         =   0   'False
         X1              =   0
         X2              =   320
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   1800
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   2
      Top             =   9480
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox impPictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3600
      Index           =   5
      Left            =   5280
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   8
      Top             =   5760
      Width           =   4800
      Begin VB.Line scanLine 
         BorderColor     =   &H0000FF00&
         Index           =   5
         Visible         =   0   'False
         X1              =   0
         X2              =   320
         Y1              =   0
         Y2              =   0
      End
   End
End
Attribute VB_Name = "frmSLIDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim picColor(0 To 2, 0 To 1000, 0 To 1000) As Integer
Dim xWidth As Long
Dim yHeight As Long
Dim isTimer As Boolean
Dim memDC As Long
Dim pHBitmap As Long


Private Sub cmdClose_Click()
   If isTimer Then
      Timer1.Enabled = False
   End If
    Unload Me
End Sub

Private Sub cmdDefault_Click()
lblStatus.Caption = "Processing ........"
 lblStatus.Refresh
 InitSlider
 lblStatus.Caption = "Processing End........"
 lblStatus.Refresh
    
End Sub

Private Sub cmdSharpen_Click()
    
    lblStatus.Caption = "Processing ........"
    lblStatus.Refresh
    
        
        Dim i As Integer
        Dim j As Integer
        Dim R As Integer
        Dim G As Integer
        Dim b As Integer
        Dim iPro As Long
        Dim k As Long
        Pro.Visible = True
        
        For k = 0 To 7
            
            If GetPixel(k) = False Then
                Exit Sub
            End If
            
           
            iPro = 0
            
            Pro.Min = 0
            Pro.Max = CLng(xWidth) * CLng(yHeight)
            scanLine(k).Visible = True
            For i = 1 To yHeight - 2
                For j = 1 To xWidth - 2
                    R = picColor(0, i, j) + 0.5 * (picColor(0, i, j) - picColor(0, i - 1, j - 1))
                    G = picColor(1, i, j) + 0.5 * (picColor(1, i, j) - picColor(1, i - 1, j - 1))
                    b = picColor(2, i, j) + 0.5 * (picColor(2, i, j) - picColor(2, i - 1, j - 1))
                    If R > 255 Then R = 255
                    If R < 0 Then R = 0
                    If G > 255 Then G = 255
                    If G < 0 Then G = 0
                    If b > 255 Then b = 255
                    If b < 0 Then b = 0
                    impPictureBox(k).PSet (j, i), RGB(R, G, b)
                   
                    Pro.Value = iPro
                    iPro = iPro + 1
                Next j
                
                scanLine(k).X1 = 0
                scanLine(k).Y1 = i + 2
                scanLine(k).X2 = 320
                scanLine(k).Y2 = i + 2
                scanLine(k).Refresh
               
                
            Next i
            scanLine(k).Visible = False
            scanLine(k).Refresh
        Next k
        
        Pro.Visible = False
        lblStatus.Caption = "Processing End"
End Sub

Private Sub Form_Load()
    Dim selectFile As String
    cmdClose.Enabled = False
    Dim i As Long
    
    idBox.Text = selectId
    InitSlider
        
End Sub
Public Sub InitSlider()
    Dim selectFile As String
    cmdClose.Enabled = False
    Dim i As Long
    
    lblStatus.Caption = "Processing ........"
    selectFile = "image/" & selectId

    
    
    For i = 0 To 7
         
         SetBitmap2 impPictureBox(i).hdc, selectFile & Str(25 + i) & ".bmp"
         PutBitmap2 impPictureBox(i).hdc
         
         impPictureBox(i).ScaleMode = 3 ' pixels
         Call StretchBlt(impPictureBox(i).hdc, 0, impPictureBox(i).Height, impPictureBox(i).Width, impPictureBox(i).Height * -1, impPictureBox(i).hdc, 0, 0, 320, 240, SRCCOPY)
         impPictureBox(i).Refresh
    Next i
    
    cmdClose.Enabled = True
    lblStatus.Caption = "Processing End"
    
End Sub

Private Sub ImageRotation(ByVal count As Long)
'회전
    Dim R As Long
    Dim G As Long
    Dim b As Long
    
    Dim i As Integer
    Dim j As Integer
    Dim iPro As Long
    If GetPixel(count) = False Then
        Exit Sub
    End If
    
    SetpicTmpScale (count)
    
    
    iPro = 0
    
    Pro.Min = 0
    Pro.Max = CLng(xWidth) * CLng(yHeight)
    picTmp.Height = impPictureBox(count).Height
    picTmp.Width = impPictureBox(count).Width
    For i = 0 To xWidth - 1
        For j = 0 To yHeight - 1
            picTmp.PSet (i, j), RGB(picColor(0, yHeight - j - 1, i), picColor(1, yHeight - j - 1, i), picColor(2, yHeight - j - 1, i))
            Pro.Value = iPro
            iPro = iPro + 1
        
        Next j

    Next i
    
    impPictureBox(count).Picture = picTmp.Image
    impPictureBox(count).Visible = True
    
    
     
End Sub



Private Sub impPictureBox_Click(index As Integer)
    'detailForm.Picture1.Picture = impPictureBox(Index).Picture
    ImageImdex = index
    
    detailForm.Image1.Picture = impPictureBox(index).Image
    detailForm.Show 1
End Sub

Private Sub Timer1_Timer()
    cmdClose.Enabled = False
    Dim selectFile As String
    Dim i As Long
   
    
    lblStatus.Caption = "Processing ........"
    selectFile = "image/" & selectId
    
    
    For i = 0 To 7
         
         SetBitmap2 impPictureBox(i).hdc, selectFile & Str(25 + i) & ".bmp"
         PutBitmap2 impPictureBox(i).hdc
         
         impPictureBox(i).ScaleMode = 3 ' pixels
         Call StretchBlt(impPictureBox(i).hdc, 0, impPictureBox(i).Height, impPictureBox(i).Width, impPictureBox(i).Height * -1, impPictureBox(i).hdc, 0, 0, 320, 240, SRCCOPY)
         impPictureBox(i).Refresh
    Next i
   
    
    
    
   
    cmdClose.Enabled = True
    lblStatus.Caption = "Processing End"
    
End Sub



Public Sub SetBitmap2(ByVal pDC As Long, ByVal fileNameToRead As String)
    DeleteBitmap2
    pHBitmap = MakeDDBFromDIB(pDC, fileNameToRead)
    
    memDC = CreateCompatibleDC(pDC)
    SelectObject memDC, pHBitmap
          
End Sub

Public Sub PutBitmap2(ByVal hdc As Long)
     BitBlt hdc, 0, 0, 320, 240, memDC, 0, 0, SRCCOPY
    
   
     
End Sub
Public Sub DeleteBitmap2()
     
     If pHBitmap <> 0 Then
          DeleteDC memDC
          DeleteObject pHBitmap
          memDC = 0
          pHBitmap = 0
     End If
     
End Sub



Public Function GetPixel(ByVal count As Long) As Boolean
'picMain의 픽셀값을 얻는다

    Dim lMask As Long
    Dim R As Long
    Dim G As Long
    Dim b As Long
    
    Dim i As Integer
    Dim j As Integer
    Dim iPro As Long
    

    xWidth = impPictureBox(count).ScaleWidth
    yHeight = impPictureBox(count).ScaleHeight
    
    If xWidth > 1000 Or yHeight > 1000 Then
        MsgBox "그림이 너무 큽니다. 이미지 변환 작업을 할수 없습니다."
        GetPixel = False
        Exit Function
    End If
    
    For i = 0 To yHeight - 1
        For j = 0 To xWidth - 1
            lMask& = impPictureBox(count).Point(j, i)
            R = lMask& Mod 256 '빨간색 추출
            G = ((lMask& And &HFF00) / 256&) Mod 256& '초록색 추출
            b = (lMask& And &HFF0000) / 65536 '파란색 추출
            
            picColor(0, i, j) = R
            picColor(1, i, j) = G
            picColor(2, i, j) = b
        Next j
    Next i
    GetPixel = True
End Function

Public Sub SetpicTmpScale(ByVal count As Long)
'picTmp 크기를 picMain크기로 맞춘다
    picTmp.Picture = LoadPicture("")
    picTmp.Height = impPictureBox(count).Height
    picTmp.Width = impPictureBox(count).Width
    picTmp.Picture = LoadPicture
End Sub


