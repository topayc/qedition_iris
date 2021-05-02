VERSION 5.00
Begin VB.Form frmCompany 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox FileList 
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image img_Slider 
      Height          =   2175
      Left            =   4440
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ImageCount As Long
Dim CurrentIndex As Long
Dim ShowOption As Boolean


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 37
            CurrentIndex = CurrentIndex - 1
            If (CurrentIndex < 0) Then
                CurrentIndex = 0
            End If
            RenderSlider
            
        Case 39
            CurrentIndex = CurrentIndex + 1
            If (CurrentIndex >= ImageCount) Then
                CurrentIndex = ImageCount - 1
            End If
            RenderSlider
        Case 38
            CurrentIndex = 0
            RenderSlider
        Case 40
            CurrentIndex = ImageCount - 1
            RenderSlider
        Case 32
            If (ShowOption) Then
                ShowOption = False
            Else
                ShowOption = True
            End If
            RenderSlider
    End Select
End Sub

Private Sub Form_Load()
    ShowOption = True
    Width = Screen.Width '∆˚¿« ∆¯¿ª »≠∏È¿« ∆¯¿∏∑Œ ∏¬√„
    Height = Screen.Height '∆˚¿« ≥Ù¿Ã∏¶ »≠∏È¿« ≥Ù¿Ã∑Œ ∏¬√„
    BackColor = RGB(0, 0, 0)
    CurrentIndex = 0
    Left = 0
    Top = 0
    
    FileList.path = "slider"
    FileList.Pattern = "*.bmp"
    ImageCount = FileList.ListCount
    
    RenderSlider
End Sub

Public Sub RenderSlider()
    Dim fileName As String
    Dim renderWidth As Long
    Dim renderHeight As Long
    img_Slider.Picture = LoadPicture
    
    img_Slider.Stretch = False
    
    fileName = "slider/" & Trim(Str(CurrentIndex)) & ".bmp"
    img_Slider.Picture = LoadPicture(fileName)
    
    If (ShowOption) Then
        img_Slider.Left = (Screen.Width - img_Slider.Width) / 2
        img_Slider.Top = (Screen.Height - img_Slider.Height) / 2
        Exit Sub
    End If
    
    renderHeight = Screen.Height
    renderWidth = (img_Slider.Width * Screen.Height) / img_Slider.Height
    
    img_Slider.Left = (Screen.Width - renderWidth) / 2
    img_Slider.Top = 0
   
    img_Slider.Stretch = True
    img_Slider.Height = renderHeight
    img_Slider.Width = renderWidth
End Sub
