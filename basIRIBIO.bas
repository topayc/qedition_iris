Attribute VB_Name = "basIRIBIO"
Option Explicit
Option Base 0

'TMR PROTOCOL COMMAND LENGTH-----------------------------------------'
Public IRIBIO_MAIN_LOOP As Byte
Public Const IRIBIO_IDLING_MODE = 0
Public Const IRIBIO_REGIST_MODE = 1

'TMR PROTOCOL COMMAND LENGTH-----------------------------------------'
Public PROTOCOLen As Long
Public CMDTXLen As Long
Public CMDRXLen As Long

'프레임 자동 저장 및 자동 GET IRIS 관련 전역변수
Public A_SaveFrame As Long
Public A_GetIris As Long

Public showGraph As Boolean

Public FILESENDING As Boolean 'FirmWare Update
Public ImageImdex As Long
Public Ptr As Long

Public INFPKGUPsize As Long 'NREGRXLen

'Memory--------------------------------------------------------------'
Public ImgBuffer(640# * 480# - 1# + 32#) As Byte
Public RXsubCODE(5000) As Byte
Public RxCode() As Byte

'Frame 파일에 쓰기 관련----------------------------------------------'
Public Const GENERIC_WRITE As Long = &H40000000
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const CREATE_ALWAYS As Long = 2
Public fHandle As Long
Public dwWritten As Long



'UART BUFFER---------------------------------------------------------'
Public TXBuffer(0 To 31) As Byte
Public RXBuffer(0 To 31) As Byte

Public VoiceData(0 To 62000) As Byte
Public VoiceCnt As Long

'FTDI API -----------------------------------------------------------'
Public FTSTATUS As Long
Public TxBytesLen As Long
Public EventsDWord As Long
Public BytesWritten As Long
Public RxBytesLen As Long
Public BytesReturned As Long

'COMMAND MENU & FRAME PROCESSOR ====================================='
Public Const INDEX_IRISECURITY = 7

'파일저장관련 정의==================================================='
Global fs As Object
Global FOut As Object
'File system
Public FS_strINFO As String
Public INFO_TOT_SIZE As Long '4Byte NETotalSize
Public INFO_TOT_CLASS As Byte '1byte NETotalClass

'===================================================================='

'ID List 전송관련 정의==================================================='
Public IDPKGsize As Long 'NREGRXLen
Public ID_TOT_SIZE As Long '4Byte NETotalSize
Public ID_TOT_PERSON As Byte '1byte NETotalClass
Public NameList(20# * 150# - 1) As Byte
Public IDList(150# - 1) As String

'===================================================================='

Public selectId As String

Public USBCheck As Integer
Public hPbit As Long



Public Declare Function FT_Open Lib "FTD2XX.DLL" (ByVal intDeviceNumber As Integer, ByRef lngHandle As Long) As Long
Public Declare Function FT_OpenEx Lib "FTD2XX.DLL" (ByVal arg1 As String, ByVal arg2 As Long, ByRef lngHandle As Long) As Long
Public Declare Function FT_Close Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Public Declare Function FT_Read Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lpszBuffer As String, ByVal lngBufferSize As Long, ByRef lngBytesReturned As Long) As Long
Public Declare Function FT_Write Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lpszBuffer As String, ByVal lngBufferSize As Long, ByRef lngBytesWritten As Long) As Long
Public Declare Function FT_ReadAny Lib "FTD2XX.DLL" Alias "FT_Read" (ByVal lngHandle As Long, ByRef lpszBuffer As Any, ByVal lngBufferSize As Long, ByRef lngBytesReturned As Long) As Long
Public Declare Function FT_WriteAny Lib "FTD2XX.DLL" Alias "FT_Write" (ByVal lngHandle As Long, ByRef lpszBuffer As Any, ByVal lngBufferSize As Long, ByRef lngBytesWritten As Long) As Long
Public Declare Function FT_SetBaudRate Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lngBaudRate As Long) As Long
Public Declare Function FT_SetDataCharacteristics Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal byWordLength As Byte, ByVal byStopBits As Byte, ByVal byParity As Byte) As Long
Public Declare Function FT_Purge Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lngMask As Long) As Long
Public Declare Function FT_GetStatus Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lngRxBytes As Long, ByRef lngTxBytes As Long, ByRef lngEventsDWord As Long) As Long

'===========================
'CHIP ID
'===========================
Public Declare Function FTID_GetNumDevices Lib "FTChipID.DLL" (ByRef arg1 As Long) As Long
Public Declare Function FTID_GetDeviceSerialNumber Lib "FTChipID.DLL" (ByVal dwDeviceIndex As Long, ByVal lpSerialBuffer As String, ByVal BufferLength As Long) As Long
Public Declare Function FTID_GetDeviceDescription Lib "FTChipID.DLL" (ByVal dwDeviceIndex As Long, ByVal lpDescriptionBuffer As String, ByVal BufferLength As Long) As Long
Public Declare Function FTID_GetDeviceLocationID Lib "FTChipID.DLL" (ByVal dwDeviceIndex As Long, ByRef lpLocationIDBuffer As Long) As Long
Public Declare Function FTID_GetDeviceChipID Lib "FTChipID.DLL" (ByVal dwDeviceIndex As Long, ByRef lpChipIDBuffer As Long) As Long
Public Declare Function FTID_GetDllVersion Lib "FTChipID.DLL" (ByVal lpDllBuffer As String, ByVal BufferLength As Long) As Long
Public Declare Function FTID_GetErrorCodeString Lib "FTChipID.DLL" (ByVal lpLanguage As String, ByVal ErrorCode As Long, ByRef lpErrorBuffer As String, ByVal BufferLength As Long) As Long

' Return codes======================================================='
Public Const FT_OK = 0
Public Const FT_INVALID_HANDLE = 1
Public Const FT_DEVICE_NOT_FOUND = 2
Public Const FT_DEVICE_NOT_OPENED = 3
Public Const FT_IO_ERROR = 4
Public Const FT_INSUFFICIENT_RESOURCES = 5
Public Const FT_INVALID_PARAMETER = 6
Public Const FT_INVALID_BAUD_RATE = 7
Public Const FT_DEVICE_NOT_OPENED_FOR_ERASE = 8
Public Const FT_DEVICE_NOT_OPENED_FOR_WRITE = 9
Public Const FT_FAILED_TO_WRITE_DEVICE = 10
Public Const FT_EEPROM_READ_FAILED = 11
Public Const FT_EEPROM_WRITE_FAILED = 12
Public Const FT_EEPROM_ERASE_FAILED = 13
Public Const FT_EEPROM_NOT_PRESENT = 14
Public Const FT_EEPROM_NOT_PROGRAMMED = 15
Public Const FT_INVALID_ARGS = 16
Public Const FT_NOT_SUPPORTED = 17
Public Const FT_OTHER_ERROR = 18

' Word Lengths======================================================='
Public Const FT_BITS_8 = 8
Public Const FT_BITS_7 = 7
' Stop Bits=========================================================='
Public Const FT_STOP_BITS_1 = 0
Public Const FT_STOP_BITS_1_5 = 1
Public Const FT_STOP_BITS_2 = 2
' Parity============================================================='
Public Const FT_PARITY_NONE = 0
Public Const FT_PARITY_ODD = 1
Public Const FT_PARITY_EVEN = 2
Public Const FT_PARITY_MARK = 3
Public Const FT_PARITY_SPACE = 4
' Purge rx and tx buffers============================================'
Public Const FT_PURGE_RX = 1
Public Const FT_PURGE_TX = 2
'===================================================================='
Global hThread As Long
Global hThreadID As Long
Global hEvent As Long
Global EventMask As Long
Global lngHandle As Long
' Flow Control
Public Const FT_FLOW_NONE = &H0
Public Const FT_FLOW_RTS_CTS = &H100
Public Const FT_FLOW_DTR_DSR = &H200
Public Const FT_FLOW_XON_XOFF = &H400
' Modem Status
Public Const FT_MODEM_STATUS_CTS = &H10
Public Const FT_MODEM_STATUS_DSR = &H20
Public Const FT_MODEM_STATUS_RI = &H40
Public Const FT_MODEM_STATUS_DCD = &H80

Public Const FT_EVENT_RXCHAR As Long = 1
Public Const FT_EVENT_MODEM_STATUS = 2

Const WAIT_ABANDONED As Long = &H80
Const WAIT_FAILD As Long = &HFFFFFFFF
Const WAIT_OBJECT_0 As Long = &H0
Const WAIT_TIMEOUT As Long = &H102
' Flags for FT_ListDevices
Public Const FT_LIST_BY_NUMBER_ONLY = &H80000000
Public Const FT_LIST_BY_INDEX = &H40000000
Public Const FT_LIST_ALL = &H20000000
' Flags for FT_OpenEx
Public Const FT_OPEN_BY_SERIAL_NUMBER = 1
Public Const FT_OPEN_BY_DESCRIPTION = 2
'FTID ChipID
' Return codes
Public Const FTID_SUCCESS = 0
Public Const FTID_INVALID_HANDLE = 1
Public Const FTID_DEVICE_NOT_FOUND = 2
Public Const FTID_DEVICE_NOT_OPENED = 3
Public Const FTID_IO_ERROR = 4
Public Const FTID_INSUFFICIENT_RESOURCES = 5

Public Const FTID_BUFER_SIZE_TOO_SMALL = 20
Public Const FTID_PASSED_NULL_POINTER = 21
Public Const FTID_INVALID_LANGUAGE_CODE = 22
Public Const FTID_INVALID_STATUS_CODE = &HFFFFFFFF
Private Const INFINITE As Long = 1000   '&HFFFFFFFF

Public Const SRCCOPY = &HCC0020
Dim FHeader As BITMAPFILEHEADER
Dim Header As BITMAPINFO
Dim SrcPtr() As Byte
Dim xGray As Boolean
Dim xhDC As Long
Public hBitmap As Long

'BMP 파일로 저장할 프레임 갯수
Public count As Long
Public ID As String


Public Type BITMAPFILEHEADER
        bfType As Integer              ' identifier
        bfSize As Long                  ' file size
        bfReserved1 As Integer     ' reserved1
        bfReserved2 As Integer     ' reserved2
        bfOffBits As Long               ' start opset of image
End Type

Public Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
'--------------------------------------------------------------------'
Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
'--------------------------------------------------------------------'
Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(0 To 255) As RGBQUAD
End Type
'--------------------------------------------------------------------'
Public Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Const CBM_CREATEDIB = &H2
Const CBM_INIT = &H4

Const DIB_RGB_COLORS = 0
Const DIB_PAL_COLORS = 1
Const DIB_PAL_INDICES = 2
Const DIB_PAL_PHYSINDICES = 2
Const DIB_PAL_LOGINDICES = 4



Public phdc As Long

'kernel 의 파일관련  함수 선언 -----------------------------------------------
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long



'DDB -> DIB -> BMP 저장-----------------------------------------------------------
Public Declare Sub DDB2DIB Lib "DDB2DIB.dll" (ByVal hpdc As Long, ByVal hObject As Long, ByVal fileName As String)

'DDB 변환된 비트맵을 특정 dc 에 출력
Public Declare Function DrawBitmap Lib "DDB2DIB.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hbit As Long)

'DIB(BMP) 파일을 DDB 로 변환
Public Declare Function MakeDDBFromDIB Lib "DDB2DIB.dll" (ByVal hdc As Long, ByVal fileName As String) As Long

'저장된 프레임 모두 삭제
Public Declare Function DeleteAllFrame Lib "DDB2DIB.dll" (ByVal path As String)

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'****************************************
' store the result BMP file and procession
' Assign the memory and Initional
'****************************************
Public Sub Allocate(ByVal hdc As Long, ByVal Width As Integer, ByVal Height As Integer)
    
    xGray = True
    
    ' input  information about file header
    ' FHeader information size(14 byte) + Header.bmiHeader information size(40 byte)+
    ' Header.bmiColor information size(256*4 byte)
    FHeader.bfOffBits = 14 + 40 + 256 * 4 '1078
    FHeader.bfReserved1 = 0
    FHeader.bfReserved2 = 0
    ' FHeader information size(14 byte) + Header.bmiHeader information size(40 byte) +
    ' Header.bmiColor information size(256*4 byte) + ImageSize information size(300*210 = 63000 byte)
    FHeader.bfSize = 14 + 40 + 256 * 4 + CLng(Width) * CLng(Height) '66614
    FHeader.bfType = &H4D42     ' &H4D42 = 'B' + 'M' = 19778

    ' read a bitmap header .
    Header.bmiHeader.biBitCount = 8
    Header.bmiHeader.biClrImportant = 0
    Header.bmiHeader.biClrUsed = 0
    Header.bmiHeader.biCompression = 0
    Header.bmiHeader.biHeight = Height
    Header.bmiHeader.biPlanes = 1
    Header.bmiHeader.biSize = 40
    Header.bmiHeader.biSizeImage = CLng(Width) * CLng(Height)
    Header.bmiHeader.biWidth = Width
    Header.bmiHeader.biXPelsPerMeter = 3780
    Header.bmiHeader.biYPelsPerMeter = 3780
    
    'Creat  palette as 256 Gray palette
    Dim i As Long
    For i = 0 To 255
      With Header.bmiColors(i)
        .rgbRed = i
        .rgbGreen = i
        .rgbBlue = i
      End With
    Next i

End Sub


Public Sub SetBitmap(ByVal hOrgDC As Long)
    DeleteBitmap
    hBitmap = CreateDIBitmap(hOrgDC, Header.bmiHeader, CBM_INIT, ImgBuffer(0), Header, DIB_RGB_COLORS)
    xhDC = CreateCompatibleDC(hOrgDC)
    SelectObject xhDC, hBitmap
          
End Sub

Public Sub SaveCaptureFrame()
     Dim fileName As String
     If (count < 50) Then ' 8프레임만 저장
                
        fileName = "image/" & ID & Str(count) & ".bmp"
        count = count + 1
        Call DDB2DIB(phdc, hBitmap, fileName)
    End If
End Sub



Public Sub DeleteBitmap()
     
     If hBitmap <> 0 Then
          DeleteDC xhDC
          DeleteObject hBitmap
          xhDC = 0
          hBitmap = 0
     End If
     
End Sub

Public Sub PutBitmap(ByVal hdc As Long)
     BitBlt hdc, 0, 0, Header.bmiHeader.biWidth, Header.bmiHeader.biHeight, xhDC, 0, 0, SRCCOPY
End Sub


Public Sub Delay(fTime As Double)
    On Error Resume Next
     Dim StartTime As Double
  
     StartTime = Timer()
    
     Do While Timer() < StartTime + fTime
          DoEvents
     Loop
    
End Sub

Public Sub CenterForm(FormName As Form)

  FormName.Top = (Screen.Height - FormName.Height) / 2
  FormName.Left = (Screen.Width - FormName.Width) / 2
  
End Sub




