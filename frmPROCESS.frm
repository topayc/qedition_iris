VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmPROCESS 
   Caption         =   "IRIS MANAGER"
   ClientHeight    =   9930
   ClientLeft      =   3900
   ClientTop       =   345
   ClientWidth     =   14985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPROCESS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   662
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   999
   Begin VB.Frame fmePROCMENU 
      Caption         =   "IRIS SECURITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   8055
      Index           =   7
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   14775
      Begin VB.ListBox LstsecDISP 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   6615
         ItemData        =   "frmPROCESS.frx":038A
         Left            =   11640
         List            =   "frmPROCESS.frx":038C
         TabIndex        =   16
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtsecStatus 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "status"
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdNameList 
         Caption         =   "USER ID DISPLAY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         MaskColor       =   &H00FFFFC0&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtsecName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   13
         Top             =   480
         Width           =   5175
      End
      Begin VB.Frame Frame3 
         Caption         =   "COMMAND"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   6735
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   11295
         Begin VB.CheckBox chk_SaveFrame 
            Caption         =   "Check2"
            Height          =   255
            Left            =   4560
            TabIndex        =   27
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox chk_AutoGet 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1920
            TabIndex        =   26
            Top             =   480
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CommandButton cmdGetTemplate 
            Caption         =   "GET IRIS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9120
            TabIndex        =   20
            Top             =   4680
            Width           =   1815
         End
         Begin VB.CommandButton cmdPutTemplate 
            Caption         =   "PUT IRIS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9120
            TabIndex        =   19
            Top             =   4200
            Width           =   1815
         End
         Begin VB.CommandButton cmdsecCANCEL 
            Caption         =   "CANCEL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3240
            Width           =   1815
         End
         Begin VB.CommandButton cmdsecSAVE 
            Caption         =   "SAVE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9120
            MaskColor       =   &H80000002&
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3720
            Width           =   1815
         End
         Begin VB.CommandButton cmdsecDEL 
            Caption         =   "USER ID DELETE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2760
            Width           =   1815
         End
         Begin VB.PictureBox picVideo 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5355
            Left            =   240
            OLEDragMode     =   1  'Automatic
            ScaleHeight     =   375
            ScaleMode       =   0  'User
            ScaleWidth      =   500
            TabIndex        =   9
            Top             =   840
            Width           =   7575
         End
         Begin VB.Frame Frame4 
            Caption         =   "Video 500 * 375"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   9000
            TabIndex        =   6
            Top             =   720
            Width           =   2055
            Begin VB.CommandButton cmdVReg 
               Caption         =   "REGISTRATION"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   8
               Top             =   360
               Width           =   1815
            End
            Begin VB.CommandButton cmdVCertification 
               Caption         =   "CERTIFICATION"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   7
               Top             =   840
               Width           =   1815
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Common"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   9000
            TabIndex        =   21
            Top             =   2400
            Width           =   2055
            Begin VB.CommandButton cmd_About 
               Caption         =   "ABOUT US"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   23
               Top             =   3240
               Width           =   1815
            End
            Begin VB.CommandButton cmdSlider 
               Caption         =   "SLIDER"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   22
               Top             =   2760
               Width           =   1815
            End
         End
         Begin VB.Label Label2 
            Caption         =   "Delete Frame On Reg Fail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   25
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Automatically Get Iris"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   24
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Label lblsecName 
         AutoSize        =   -1  'True
         Caption         =   "NAME : "
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.Timer TMR 
      Left            =   9720
      Top             =   1080
   End
   Begin VB.Frame fmeMessage 
      Caption         =   "MESSAGE"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   14775
      Begin VB.TextBox MSGtxtResult 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   7215
      End
      Begin VB.TextBox MSGtxtHexResult 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   6975
      End
   End
   Begin MCI.MMControl mmcSOUND 
      Height          =   330
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSComDlg.CommonDialog DIALOG 
      Left            =   7080
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape shapeGREEN2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape shapeRED 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape shapeGREEN1 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblPORTSTATUS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "[PORT STATUS]"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   11040
      TabIndex        =   18
      Top             =   1200
      Width           =   1950
   End
End
Attribute VB_Name = "frmPROCESS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim flg_Capture As Boolean

Sub UsbBufClear()

     If FT_GetStatus(lngHandle, RxBytesLen, TxBytesLen, EventsDWord) <> FT_OK Then Exit Sub
     FTSTATUS = FT_ReadAny(lngHandle, ImgBuffer(0), RxBytesLen, BytesReturned)

End Sub


Public Sub Sound_Play(playname As String)
     mmcSOUND.DeviceType = ""
     mmcSOUND.Command = "Close"
     mmcSOUND.fileName = playname
     mmcSOUND.Command = "Open"
     mmcSOUND.Command = "Prev"
     mmcSOUND.Command = "Play"
End Sub


Public Sub TxBufClear()
        Dim i As Integer
        For i = 0 To 31
            TXBuffer(i) = 0
        Next i
End Sub

'================================================================================
'                         32byte Command transmission function
'================================================================================
Sub Cmd_Send(ByVal Cmd As Byte, ByVal Size As Long, ByVal Name As String, ByVal Dim27 As Byte)
                          
            Dim strLen As String
            Dim IDLen As Byte
            Dim ChkSum As Long
            Dim i, j As Integer
            
            Call TxBufClear
            
            TXBuffer(0) = &HF1
            TXBuffer(29) = &HFE
            
            TXBuffer(1) = Cmd
            TXBuffer(2) = Not (TXBuffer(1))
            
            strLen = Right$("00000000" & Hex$(Size), 8)
            TXBuffer(3) = Val("&H" + Right$(strLen, 2))
            TXBuffer(4) = Val("&H" + Mid$(strLen, 5, 2))
            TXBuffer(5) = Val("&H" + Mid$(strLen, 3, 2))
            TXBuffer(6) = Val("&H" + Left$(strLen, 2))
    
            IDLen = Len(Name)
    
            j = 7

            If IDLen <> 0 Then
                For i = 1 To IDLen
                     TXBuffer(j) = Asc(Mid$(Name, i, 1))
                     j = j + 1
                Next i
            End If
                        
            TXBuffer(27) = Dim27
            
            ChkSum = 0
            For i = 1 To 27
                ChkSum = ChkSum + TXBuffer(i)
            Next i
            
            TXBuffer(28) = ChkSum Mod 256
            
            FTSTATUS = FT_WriteAny(lngHandle, TXBuffer(0), 32, BytesWritten)
                                   
End Sub



Private Sub chk_AutoGet_Click()
    A_GetIris = chk_AutoGet.Value
End Sub

Private Sub chk_SaveFrame_Click()
    A_SaveFrame = chk_SaveFrame.Value
End Sub



Private Sub cmd_About_Click()
     frmCompany.Show 1
End Sub

Private Sub cmdGetTemplate_Click()

    Dim ID As String

    MSGtxtResult.Text = ""
    ID = txtsecName
    Call Cmd_Send(&HB1, 32, ID, 0)

End Sub


Private Sub cmdNameList_Click()

     Dim ID As String
    
     PROTOCOLen = CMDRXLen
     LstsecDISP.Clear

     Call Cmd_Send(6, 32, "", 0)
End Sub


Private Sub cmdPutTemplate_Click()
        Dim cnt, i As Long
        Dim Name, Info As String
        Dim fBin, bTmp As Byte
        Dim fileName As String
        
        DIALOG.CancelError = False
        DIALOG.Flags = cdlOFNHideReadOnly
        DIALOG.InitDir = App.path
        DIALOG.Filter = "Binary(BIN)|*.bin"
        DIALOG.FilterIndex = 1
        DIALOG.ShowOpen
            
        
        fileName = DIALOG.fileName
        
        If (Len(fileName) = 0 Or fileName = " ") Then
            MsgBox "홍채 인식기 모듈로 전송하고자 하는 홍채템플릿을 선택해주세요"
            Exit Sub
        End If
        
    
        Name = ""
        Info = ""
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
                
        Dim SERIALLen As Byte
        Dim ChkSum As Double
        Dim StrCmdLen As String

        Call TxBufClear
        
        TXBuffer(0) = &HF1
        TXBuffer(1) = &HB4
        TXBuffer(2) = Not (TXBuffer(1))
            
        StrCmdLen = Right$("00000000" & Hex$(INFO_TOT_SIZE), 8)         'file size 4 byte
        TXBuffer(3) = Val("&H" + Right$(StrCmdLen, 2))
        TXBuffer(4) = Val("&H" + Mid$(StrCmdLen, 5, 2))
        TXBuffer(5) = Val("&H" + Mid$(StrCmdLen, 3, 2))
        TXBuffer(6) = Val("&H" + Left$(StrCmdLen, 2))
            
        SERIALLen = Len(Name)
        i = 7
        For cnt = 1 To SERIALLen
                TXBuffer(i) = Asc(Mid(Name, cnt, 1))
                i = i + 1
        Next cnt
        
        TXBuffer(27) = 0
            
        ChkSum = 0
        For cnt = 1 To 27
                ChkSum = ChkSum + TXBuffer(cnt)
        Next cnt
        
        TXBuffer(28) = ChkSum Mod 256
        TXBuffer(29) = &HFE
            
        ReDim RxCode(0 To INFO_TOT_SIZE - 1) As Byte
        For cnt = 0 To INFO_TOT_SIZE - 1
            Get fBin, , bTmp
            RxCode(cnt) = bTmp
        Next cnt

        Close fBin
        
        BytesWritten = 0
        
        FTSTATUS = FT_WriteAny(lngHandle, TXBuffer(0), 32, BytesWritten)

        If FTSTATUS <> FT_OK Then
             MSGtxtResult.Text = "FT_WriteAny(REGISTRATION) Failed status=" & FTSTATUS
             Exit Sub
        End If
        
        
        txtsecName = ""

End Sub

Private Sub cmdsecCANCEL_Click()
     PROTOCOLen = CMDRXLen
     flg_Capture = False
     Call Cmd_Send(4, 32, txtsecName, 0)

End Sub


Private Sub cmdsecCERTI_Click()
     
     Dim ID As String
    
     PROTOCOLen = CMDRXLen              '32 Byte Command Mode

     ID = txtsecName
     MSGtxtResult = "Please look at the mirror for certification."
     Call Cmd_Send(2, 32, ID, 0)
     If (FTSTATUS = FT_OK) Then Call Sound_Play(App.path + "\identify_ready.wav")
    
End Sub


Private Sub cmdsecDEL_Click()
     
     Dim ID As String
     Dim response

     ID = txtsecName

     If (Len(txtsecName) <> 0) Then
          response = MsgBox("We will delete ID[ " + MSGtxtResult.Text + " ]  sure?", vbYesNo + vbCritical + vbDefaultButton2, "ID Delete")
          MSGtxtResult.Text = txtsecName + "  will be deleted."
          If response <> vbYes Then
               MSGtxtResult.Text = ""
               Exit Sub
          Else
               Call Cmd_Send(3, 32, ID, 0)
          End If
     Else
          response = MsgBox("WE will delete all ID, are you sure ?", vbYesNo + vbCritical + vbDefaultButton2, "ALL DELETE")
          MSGtxtResult.Text = "We will delete all ID."
          If response = vbYes Then
               response = MsgBox("Are you sure ?", vbYesNo + vbCritical + vbDefaultButton2, "ALL DELETE")
               If response <> vbYes Then
                    MSGtxtResult.Text = ""
                    Exit Sub
               Else
                    Call Cmd_Send(3, 32, "", 0)
               End If
          Else
               MSGtxtResult.Text = ""
               Exit Sub
          End If
    End If
    
End Sub



Private Sub cmdsecREGIST_Click()
     
     Dim ID As String

     PROTOCOLen = CMDRXLen              '32 Byte Command Mode

     ID = txtsecName
     
     MSGtxtResult = "Please look at the mirror for registration."
        
     IRIBIO_MAIN_LOOP = IRIBIO_REGIST_MODE
     
     Call Cmd_Send(1, 32, ID, 0)
     
     If (FTSTATUS = FT_OK) Then Call Sound_Play(App.path + "\regi_ready.wav")
         
End Sub



Private Sub cmdsecSAVE_Click()
     
     PROTOCOLen = CMDRXLen
     Call Cmd_Send(5, 32, "", 0)
    
End Sub



Sub ReadUsb()

     Do
          DoEvents
          EventsDWord = 0
          RxBytesLen = 0
          
          If (FT_GetStatus(lngHandle, RxBytesLen, TxBytesLen, EventsDWord) <> FT_OK) Then Exit Sub
          
     Loop While RxBytesLen < 32
                                                                                                                                        
     BytesReturned = 0
     
     FTSTATUS = FT_ReadAny(lngHandle, RXBuffer(0), 32, BytesReturned)
 
End Sub


Sub UsbBuffClear()
      
     If FT_GetStatus(lngHandle, RxBytesLen, TxBytesLen, EventsDWord) <> FT_OK Then Exit Sub
        
     If RxBytesLen > 0 Then
          FTSTATUS = FT_ReadAny(lngHandle, ImgBuffer(0), RxBytesLen, BytesReturned)
     End If
        
End Sub



Private Sub cmdSlider_Click()
     Dim ID As String
     Dim response
     Dim selectFile As String

     ID = txtsecName
     selectId = ID

     If (Len(txtsecName) <= 0) Then
        response = MsgBox("슬라이더로 보고자 하는 ID 를 선택해주세요 ")
        Exit Sub
     End If
    
     selectFile = "image/" & ID & Str(0) & ".bmp"
     If Dir(selectFile) = "" Then

      response = MsgBox("아이디가 [ " + ID + " ] 인 홍채 이미지가 존재하지 않습니다" & Chr(13) & "아이디를 확인해주세요")
       Exit Sub
    End If


     
     frmSLIDER.Show 1
    


End Sub

Private Sub cmdVCertification_Click()
     
     On Error GoTo ErrorHandler
     
     Dim ID As String
     Dim pAddr, pDiff As Long
     Dim BytesReturned, result As Long
     Dim i As Integer
     Dim count As Long
     Dim fileName As String
     count = 0
                
     TMR.Enabled = False
     
     MSGtxtResult = "Please look at the mirror for certification."
                     
     Call UsbBuffClear
     Call Cmd_Send(2, 32, txtsecName, 1)
                    
     If (FTSTATUS = FT_OK) Then Call Sound_Play(App.path + "\identify_ready.wav")
                     
     Allocate picVideo.hdc, 320, 240
                       
     Call DeleteBitmap
     flg_Capture = True
     pAddr = 0
     ID = ""
     
     Do
          Call ReadUsb
                            
          If (RXBuffer(0) = &HF4 And RXBuffer(29) = &HFE) Then
                                
               Select Case RXBuffer(1)
                                          
                    Case &HC0
                    
                         Do                                      'Iris High Image
                              DoEvents
                                                
                              If FT_GetStatus(lngHandle, RxBytesLen, TxBytesLen, EventsDWord) <> FT_OK Then Exit Do
                                   
                              If flg_Capture = False Then Exit Do
                                            
                         Loop While RxBytesLen < 38400
                                                                     
                                                                                                          
                         FTSTATUS = FT_ReadAny(lngHandle, ImgBuffer(0), RxBytesLen, BytesReturned)
                                    
                         pDiff = 76800 - RxBytesLen
                         pAddr = RxBytesLen

                         RxBytesLen = 0
                         
                         
                         Do                                      'Iris Low Image
                              DoEvents
                                                
                              If FT_GetStatus(lngHandle, RxBytesLen, TxBytesLen, EventsDWord) <> FT_OK Then Exit Sub
                                                
                              If flg_Capture = False Then Exit Do
                              
                         Loop While RxBytesLen < pDiff
                                                                                        
                         FTSTATUS = FT_ReadAny(lngHandle, ImgBuffer(pAddr), pDiff, BytesReturned)
                                            
                         SetBitmap picVideo.hdc
                         PutBitmap picVideo.hdc
                         picVideo.ScaleMode = 3 ' pixels
                         Call StretchBlt(picVideo.hdc, 0, picVideo.ScaleHeight, picVideo.ScaleWidth, picVideo.ScaleHeight * -1, picVideo.hdc, 0, 0, 320, 240, SRCCOPY)
                          
                         picVideo.Refresh             'Image Display
                                                                                
                    Case &H2                             'CERTIFI Fail
                         Call Sound_Play(App.path + "\IDFAIL.wav")
                                   
                         MSGtxtResult.Text = "  is not certified!"
                         TMR.Enabled = True
                         Exit Do
                    Case &H14                          'Err=-1 Canceled
                         MSGtxtResult.Text = "Canceled..."
                         Call Sound_Play(App.path + "\CANCEL.wav")
                         TMR.Enabled = True
                         Exit Do
                                                                                                       
                    Case &HF2                          'Success to Certification
                         ID = ""
                         For i = 7 To 27
                              If RXBuffer(i) > &H20 And RXBuffer(i) < &H7F Then
                                   ID = ID & Chr(RXBuffer(i))
                              End If
                         Next i
                         
                         MSGtxtResult.Text = ID & "  is certified!"
                         Call Sound_Play(App.path + "\IDOK.wav")
                         TMR.Enabled = True
                         Exit Do
                   Case &H21, &H22
                         Call Sound_Play(App.path + "\ding.wav")
                         MSGtxtResult.Text = " No frames captured..."
                         Exit Do
                   Case &H32                                                              '-------------- Err=-3 Database is Empty
                         Call Sound_Play(App.path + "\ding.wav")
                         MSGtxtResult.Text = " Database is empty..."
                         
                    Case &H77
                         Call Sound_Play(App.path + "\ding.wav")
                         MSGtxtResult.Text = " System Init !"
                         Exit Do
                         
                         
                    Case &H79
                         Call Sound_Play(App.path + "\TIMEOVER.wav")
                         MSGtxtResult.Text = " Time Over !"
                         Exit Do
                         
                         
                    Case &H31
                         MSGtxtResult.Text = "ID Name Error !"
                         Call Sound_Play(App.path + "\chord.wav")
                         Exit Do
                         
                    Case Else
                         TMR.Enabled = True
                         MSGtxtResult.Text = "ERROR Number " & RXBuffer(1)
                         Call Sound_Play(App.path + "\ding.wav")
                         
                         Exit Do
                                          
               End Select
               
          End If
          
     Loop While True
                                             
     Call UsbBuffClear                            'USB Buffer Clear
                                            

ErrorHandler:
         TMR.Enabled = True

End Sub

Private Sub cmdVReg_Click()
     
     On Error GoTo ErrorHandler
     Dim count As Long
    
     Dim pAddr, pDiff As Long
     Dim BytesReturned, result As Long
     Dim i As Integer
     Dim fileName As String
     count = 0
     ID = txtsecName
     
     If (Len(ID) <= 0 Or ID = "") Then
        MsgBox "등록하시고자 하는 ID 를 먼저 입력해 주셔야 합니다"
        Exit Sub
     End If
     
     phdc = picVideo.hdc
     
     TMR.Enabled = False                          'Timer Stop
     
     MSGtxtResult = "Please look at the mirror for registration."
     Call UsbBuffClear
        
     IRIBIO_MAIN_LOOP = IRIBIO_REGIST_MODE
     
     Call Cmd_Send(1, 32, txtsecName, 1)                 '(Video Mode) Register Command

     Call Sound_Play(App.path + "\regi_ready.wav")
                                                               
     Allocate picVideo.hdc, 320, 240
                       
     Call DeleteBitmap
     
     flg_Capture = True
     pAddr = 0
     
     Do
          Call ReadUsb
                            
          If (RXBuffer(0) = &HF4 And RXBuffer(29) = &HFE) Then
                
           
                                
               Select Case RXBuffer(1)
                                          
                    Case &HC0
                    
                         Do                                 'Iris High Image
                              DoEvents
                                           
                              If FT_GetStatus(lngHandle, RxBytesLen, TxBytesLen, EventsDWord) <> FT_OK Then Exit Sub
                                                                                                            
                              If flg_Capture = False Then Exit Do
                                       
                         Loop While RxBytesLen < 38400
                                                                                                                                                          
                         FTSTATUS = FT_ReadAny(lngHandle, ImgBuffer(0), RxBytesLen, BytesReturned)
                               
                         pDiff = 76800 - RxBytesLen
                         pAddr = RxBytesLen
                                       
                         RxBytesLen = 0
                    
                         Do                                 'Iris Low Image
                              DoEvents
                                           
                              If FT_GetStatus(lngHandle, RxBytesLen, TxBytesLen, EventsDWord) <> FT_OK Then Exit Sub
                                           
                              If flg_Capture = False Then Exit Do
                         
                         Loop While RxBytesLen < pDiff
                                               
                                               
                         FTSTATUS = FT_ReadAny(lngHandle, ImgBuffer(pAddr), pDiff, BytesReturned)
                                       
                         SetBitmap picVideo.hdc
                         SaveCaptureFrame
                         PutBitmap picVideo.hdc
                         picVideo.ScaleMode = 3 ' pixels
                         Call StretchBlt(picVideo.hdc, 0, picVideo.ScaleHeight, picVideo.ScaleWidth, picVideo.ScaleHeight * -1, picVideo.hdc, 0, 0, 320, 240, SRCCOPY)
                         picVideo.Refresh            'Image Display
                         
                         
                         
                         
                         
                 
                    Case &H1                           'REGIST Fail
                         Call Sound_Play(App.path + "\REGFAIL.wav")
                         MSGtxtResult.Text = " is not registered !" + " Class=" + Hex(RXBuffer(27))
                         
                         ' 등록 실패시 옵션에 따라 저장된 모든 프레임 삭제
                         If (A_SaveFrame = 1) Then
                            DeleteAllFrame ("image/" & txtsecName & "*.bmp")
                         End If
                         
                         Exit Do
                    Case &H14                          ' Canceled
                         MSGtxtResult.Text = "Canceled..."
                         Call Sound_Play(App.path + "\CANCEL.wav")
                         
                         ' 등록 취소 옵션에 따라 저장된 모든 프레임 삭제
                         If (A_SaveFrame = 1) Then
                            DeleteAllFrame ("image/" & txtsecName & "*.bmp")
                         End If
                         
                         Exit Do
                    Case &HF1                          'Success to Regist
                         ID = ""
                         count = 0
                          
                         For i = 7 To 27
                              If RXBuffer(i) > &H20 And RXBuffer(i) < &H7F Then
                                   ID = ID & Chr(RXBuffer(i))
                              End If
                         Next i
                         
                         MSGtxtResult.Text = ID & " is registered !" + " Class=" + Hex(RXBuffer(27))
                         Call Sound_Play(App.path + "\REGOK.wav")
                         
                         If (A_GetIris = 1) Then
                           Call Cmd_Send(&HB1, 32, ID, 0) ' 등록 성공시 자동 GET IRIS
                         End If
                         
                         Exit Do
                         
                    Case &H21, &H22
                         Call Sound_Play(App.path + "\ding.wav")
                         MSGtxtResult.Text = " No frames captured..."
                         Exit Do
                         
                    Case &H77
                         Call Sound_Play(App.path + "\ding.wav")
                         MSGtxtResult.Text = " System Init !"
                         Exit Do
                         
                         
                    Case &H79
                         Call Sound_Play(App.path + "\TIMEOVER.wav")
                         MSGtxtResult.Text = " Time Over !"
                         Exit Do
                    Case &H51                                                                            '-------------- Err=-5 Not enough frames for regist
                         Call Sound_Play(App.path + "\ding.wav")
                         MSGtxtResult.Text = " Not enough frames for registration"
                         Exit Do
                    Case &H31
                         MSGtxtResult.Text = "ID Name Error"
                         Call Sound_Play(App.path + "\chord.wav")
                         Exit Do
                         
                    Case Else
                         MSGtxtResult.Text = "ERROR Number " & RXBuffer(1)
                         Call Sound_Play(App.path + "\ding.wav")
                         
                         Exit Do
                                          
               End Select
                                    
          End If
                                                                                            
     Loop While True
     
     TMR.Enabled = True
          
     Call UsbBuffClear                                                          'buffer clear
          
ErrorHandler:
         TMR.Enabled = True

End Sub



Private Sub Form_Load()
    CenterForm Me '폼을 화면 가운데로
    CMDTXLen = 32
    CMDRXLen = 32

    INFPKGUPsize = 0
    
     fmeMessage.Caption = "MESSAGE"
     lblPORTSTATUS.Caption = "[PORT CLOSE]"
     
     shapeGREEN1.Visible = False
     shapeGREEN2.Visible = False
     shapeRED.Visible = True
    
     fmeMessage.Enabled = False
     chk_AutoGet.Value = 0

     FTSTATUS = FT_Open(0, lngHandle)
        
     If FTSTATUS <> FT_OK Then
          MSGtxtResult.Text = ""
          MSGtxtHexResult.Text = "FT_SetDataCharacteristics() Failed status=" & FTSTATUS
          Exit Sub
     Else
          MSGtxtHexResult.Text = "SUCCESS to DRIVER USB Port OPEN"""
          FTSTATUS = FT_Purge(lngHandle, FT_PURGE_RX Or FT_PURGE_TX)

          PROTOCOLen = CMDRXLen
          TMR.Interval = 60
          shapeGREEN1.Visible = True
          shapeGREEN2.Visible = False
          shapeRED.Visible = False
          PROTOCOLen = CMDRXLen
     End If
  
End Sub


Private Sub LED_DISPLAY()
     If (shapeGREEN1.Visible) Then
          shapeGREEN2.Visible = True
          shapeRED.Visible = False
          shapeGREEN1.Visible = False
     ElseIf (shapeGREEN2.Visible) Then
          shapeGREEN1.Visible = True
          shapeRED.Visible = False
          shapeGREEN2.Visible = False
     End If
End Sub


Private Sub Form_Terminate()
     
     If FT_Close(lngHandle) <> FT_OK Then
          MSGtxtResult.Text = ""
          MSGtxtHexResult.Text = "FT_Close() Close Failed"
          Exit Sub
     Else
          MSGtxtHexResult.Text = "FT_Close() Close OK"
          TMR.Interval = 0
     End If
     
     End

End Sub



Private Sub LstsecDISP_Click()
    txtsecName.Text = LstsecDISP.List(LstsecDISP.ListIndex)
End Sub



Private Sub TMR_Timer()
    Dim strData() As Byte
     Dim bArray() As Byte
        
     Dim i, j, iDcnt, cnt As Long
     Dim fBin As Integer
     Dim fileName As String

     Dim IRIBIOSTRING, IDName As String
     Dim CTemp As Byte
        
     Call LED_DISPLAY
    
     If (PROTOCOLen = CMDRXLen) Then
          
          EventsDWord = 0
          RxBytesLen = 0
          
          FTSTATUS = FT_GetStatus(lngHandle, RxBytesLen, TxBytesLen, EventsDWord)
          
          If (FTSTATUS = FT_OK) Then
               If (RxBytesLen = CMDRXLen) Then
                    BytesReturned = 0
                    FTSTATUS = FT_ReadAny(lngHandle, RXBuffer(0), 32, BytesReturned)
                    
                    If (RXBuffer(0) = &HF4 And RXBuffer(29) = &HFE) Then                       'DSP->PC
                         IDName = ""
                         
                          For i = 7 To 27
                                If RXBuffer(i) > &H20 And RXBuffer(i) < &H7F Then
                                        IDName = IDName & Chr(RXBuffer(i))
                                End If
                         Next i
                                     
                         Select Case RXBuffer(1)
                              
                              Case &H1                                                                               '-------------REGIST Fail err=0
                                   Call Sound_Play(App.path + "\REGFAIL.wav")
                                   MSGtxtResult.Text = IDName & " is not registered !" + " Class=" + Hex(RXBuffer(27))
                              
                              Case &H2                                                                               '------------CERTIFI Fail err=0
                                   Call Sound_Play(App.path + "\IDFAIL.wav")
                                   MSGtxtResult.Text = IDName & "  is not certified!"
                              
                              Case &H3                                                                               '-------------Delete Error
                                   Call Sound_Play(App.path + "\chord.wav")
                                   MSGtxtResult.Text = IDName & "  was not found!"
                              
                              Case &H6                                                                               '------------ FileList
                                   If (RXBuffer(27) = 0) Then
                                       MSGtxtResult.Text = ""
                                       MSGtxtResult.Text = " DataBase Empty!!!"
                                       ID_TOT_PERSON = 0
                                   
                                   Else
                                       MSGtxtResult.Text = ""
                                       MSGtxtResult.Text = "  Transmission..."
                                       IDPKGsize = Val("&H" & Right$("00" & Hex$(RXBuffer(6)), 2) & Right$("00" & Hex$(RXBuffer(5)), 2) & Right$("00" & Hex$(RXBuffer(4)), 2) & Right$("00" & Hex$(RXBuffer(3)), 2))
                                       PROTOCOLen = IDPKGsize
                                    
                                       ID_TOT_SIZE = IDPKGsize
                                       ID_TOT_PERSON = RXBuffer(27)
                                   End If
                                   
                                   'txtsecStatus(3).Text = "0 /" + Str(RXBuffer(27))
                                   txtsecStatus(3).Text = Str(ID_TOT_PERSON)
                                                    
                              Case &H10                                                                              '--------------Version Display
                                   For i = 0 To 19
                                        IRIBIOSTRING = IRIBIOSTRING + Chr(RXBuffer(i + 7))
                                   Next i
                              
                              Case &H14                                                                              '---------------Err=-1 Canceled
                                   MSGtxtResult.Text = "Canceled..."
                                   Call Sound_Play(App.path + "\chimes.wav")
                                   
                                              
                              Case &H22, &H21                                                                   '-------------Err=-2 No frames captured
                                   Call Sound_Play(App.path + "\ding.wav")
                                   MSGtxtResult.Text = " No frames captured..."
                                   
                              Case &H32, &H31                                                                 '-------------- Err=-3 Database is Empty
                                   Call Sound_Play(App.path + "\ding.wav")
                                   MSGtxtResult.Text = " Database is empty..."
                                   
                              Case &H42, &H41                                                                 '-------------- Err=-4 Database is full
                                   Call Sound_Play(App.path + "\ding.wav")
                                   MSGtxtResult.Text = " Database is full..."
                        
                              Case &H51                                                                            '-------------- Err=-5 Not enough frames for regist
                                   Call Sound_Play(App.path + "\ding.wav")
                                   MSGtxtResult.Text = " Not enough frames for registration"
                    
                              Case &H65                                                                            '-------------- Err=-6 Failed to save
                                   Call Sound_Play(App.path + "\ding.wav")
                                   MSGtxtResult.Text = " Failed... Pls save again."
                                                          
                              Case &H79
                                    Call Sound_Play(App.path + "\TIMEOVER.wav")
                                    MSGtxtResult.Text = " Time Over !"
                                                
                              Case &H9A                                                                       '---------------Success download
                                   Call Sound_Play(App.path + "\chimes.wav")
                                   MSGtxtResult.Text = "Success to Download"
                                                                                                                                    
                              Case &H9F                                                                       'checksum error
                                   Call Sound_Play(App.path + "\ding.wav")
                                   MSGtxtResult.Text = "Checksum ERROR"
                         
                              Case &HB1                                                                       '--Network transmission'0xB1 을 수신 받은 후 바로 Info.Pkg. Infomation이 전송되어진다.
                                   MSGtxtResult.Text = "  Transmission..."
                                   INFPKGUPsize = Val("&H" & Right$("00" & Hex$(RXBuffer(6)), 2) & Right$("00" & Hex$(RXBuffer(5)), 2) & Right$("00" & Hex$(RXBuffer(4)), 2) & Right$("00" & Hex$(RXBuffer(3)), 2))
                                   PROTOCOLen = INFPKGUPsize
                              
                                   
                                   INFO_TOT_SIZE = INFPKGUPsize
                                   INFO_TOT_CLASS = RXBuffer(27)
                                            
                              Case &HB4
                                   MSGtxtResult.Text = IDName & "   Regist Code download complete"
                                   Call Sound_Play(App.path + "\chimes.wav")
                                   Call cmdNameList_Click
                              
                              Case &HB5
                                   MSGtxtResult.Text = "INFOPKGFORM ERROR  (DOWNLOAD)"
                                   Call Sound_Play(App.path + "\chord.wav")
                                                           
                              Case &HB7
                                   MSGtxtResult.Text = "ERROR, FILENAME DO NOT EXIST "
                                   Call Sound_Play(App.path + "\chord.wav")
                              
                              Case &HB8
                                    MSGtxtResult.Text = "ERROR, FILENAME EXIST"
                                    Call Sound_Play(App.path + "\chord.wav")
                                    
                              Case &HB9
                                    MSGtxtResult.Text = "ERROR, DATABASE IS FULL (DOWNLOAD)"
                                    Call Sound_Play(App.path + "\chord.wav")
                              
                              Case &HBA
                                    MSGtxtResult.Text = "Possible to DOWNLOAD"
                                    
                                    FTSTATUS = FT_WriteAny(lngHandle, RxCode(0), INFO_TOT_SIZE, BytesWritten)
                              
                                    If FTSTATUS <> FT_OK Then
                                         MSGtxtResult.Text = "FT_WriteAny(REGISTRATION) Failed status=" & FTSTATUS
                                         Exit Sub
                                    End If
                                                                     
                                               
                              Case &HF1                                              'Success to Regist Err=1
                                   MSGtxtResult = IDName & " is registered !" + " Class=" + Hex(RXBuffer(27))
                                   Call Sound_Play(App.path + "\REGOK.wav")
                        
                              Case &HF2                                              'Success to Certification Err=1
                                   MSGtxtResult = IDName & "  is certified!"
                                   Call Sound_Play(App.path + "\IDOK.wav")
                        
                              Case &HF3                                              'Success to Delete Err=1
                              
                                   If (txtsecName.Text = "") Then
                                        MSGtxtResult.Text = " All deleted!"
                                   Else
                                        MSGtxtResult = IDName & " is deleted."
                                   End If
                                   Call Sound_Play(App.path + "\ding.wav")
                                   Call cmdNameList_Click
                        
                              Case &HF5                                              'Success to save Err=7
                                   MSGtxtResult.Text = " Saved !"
                                   Call Sound_Play(App.path + "\chimes.wav")
                                                              
                              Case Else
                                   FILESENDING = False
                         End Select
                    End If
               Else
                    If (RxBytesLen > 0 And RxBytesLen <> CMDRXLen) Then FTSTATUS = FT_Purge(lngHandle, FT_PURGE_RX)
               End If
               
          End If
             
     ElseIf (PROTOCOLen = INFPKGUPsize) Then
          
          EventsDWord = 0
          FTSTATUS = FT_GetStatus(lngHandle, RxBytesLen, TxBytesLen, EventsDWord)
          If (FTSTATUS = FT_OK) Then
               If (RxBytesLen = INFPKGUPsize) Then
                    'ReDim RXsubCODE(0 To INFPKGUPsize - 1) As Byte
                    BytesReturned = RxBytesLen = 0
                    FTSTATUS = FT_ReadAny(lngHandle, RXsubCODE(0), INFPKGUPsize, BytesReturned)
                
                    FS_strINFO = "QIRIBIOM60622-WINWIN"
                
                    frmPROCESS.MousePointer = 11
                    fileName = "REG_" + txtsecName + ".bin"
            
                    fBin = FreeFile()
                    Open fileName For Binary Access Write As fBin
                    'Header informatin write------------------------------------------------
                    For i = 1 To 20
                         CTemp = Asc(Mid(FS_strINFO, i, 1))
                         Put fBin, , CTemp
                    Next i
                    'Header IDName----------------------------------------------------------

                    For i = 1 To 20
                         If (Mid(txtsecName, i, 1) = "") Then
                              CTemp = 0
                         Else
                              CTemp = Asc(Mid(txtsecName, i, 1))
                         End If

                         Put fBin, , CTemp
                    Next i

                    Put fBin, , INFO_TOT_SIZE
                    Put fBin, , INFO_TOT_CLASS
    
                    For i = 0 To UBound(RXsubCODE)
                         CTemp = RXsubCODE(i)
                         Put fBin, , CTemp
                    Next i
                    
                    Close fBin
                    
                    MSGtxtResult.Text = "REG_" + txtsecName + ".bin saved"
    
                    Call Sound_Play(App.path + "\chimes.wav")
                
                    MSGtxtResult.Text = "TRANSMISSION DONE"
                    frmPROCESS.MousePointer = 1
                    PROTOCOLen = CMDRXLen
               Else
                    If (RxBytesLen > INFPKGUPsize) Then
                         Debug.Print "FT_Purge[INFPKGUPsize]RxBytes="; Str(RxBytesLen)
                         FTSTATUS = FT_Purge(lngHandle, FT_PURGE_RX)
                         DoEvents
                    End If
               End If
          
          End If
        
        ElseIf (PROTOCOLen = IDPKGsize) Then
        USBCheck = 0
        
        EventsDWord = 0
        FTSTATUS = FT_GetStatus(lngHandle, RxBytesLen, TxBytesLen, EventsDWord)
        
        If (FTSTATUS = FT_OK) Then
            If (RxBytesLen = IDPKGsize) Then
                
                BytesReturned = RxBytesLen = 0
                FTSTATUS = FT_ReadAny(lngHandle, NameList(0), IDPKGsize, BytesReturned)
'                frmPROCESS.MousePointer = 11
                
                '========================================================================
                '   IDPKGsize 만큼 받은 것을 20Byte 단위로 나눠서 처리하는 루틴...
                '
                '   IDPKGsize       : 총 ID List 크기
                '   ID_TOT_PERSON   : 총 ID Count
                '========================================================================
                For iDcnt = 0 To (ID_TOT_PERSON - 1)
                    For cnt = 0 To 19
                        IDList(iDcnt) = IDList(iDcnt) + Chr(NameList((iDcnt * 20) + cnt))
                    Next cnt
                Next iDcnt
                
                For iDcnt = 0 To (ID_TOT_PERSON - 1)
                    LstsecDISP.AddItem IDList(iDcnt)
                Next iDcnt
                
                txtsecStatus(3).Text = Str(ID_TOT_PERSON)
                
                Call Sound_Play(App.path + "\chimes.wav")
                
                MSGtxtResult.Text = "TRANSMISSION DONE"
                frmPROCESS.MousePointer = 1
                PROTOCOLen = CMDRXLen
                DoEvents
            Else
                If (RxBytesLen > IDPKGsize) Then
                    Debug.Print "FT_Purge[IDPKGsize]RxBytes="; Str(RxBytesLen)
                    FTSTATUS = FT_Purge(lngHandle, FT_PURGE_RX)
                    DoEvents
                End If
            End If
        
        End If
     
     End If
     
     
     
End Sub
