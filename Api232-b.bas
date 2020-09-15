Attribute VB_Name = "api_Module"
'api-232.bas   use for  all  dll decalre
'  作者：  朱国魂 ------桂林电子工业学院
Global Const B50 = &H0
Global Const B75 = &H1
Global Const B110 = &H2
Global Const B134 = &H3
Global Const B150 = &H4
Global Const B300 = &H5
Global Const B600 = &H6
Global Const B1200 = &H7
Global Const B1800 = &H8
Global Const B2400 = &H9
Global Const B4800 = &HA
Global Const B7200 = &HB
Global Const B9600 = &HC
Global Const B19200 = &HD
Global Const B38400 = &HE
Global Const B57600 = &HF
Global Const B115200 = &H10
Global Const B230400 = &H11
Global Const B460800 = &H12
Global Const B921600 = &H13

'       MODE setting
Global Const BIT_5 = &H0                             ' Word length define
Global Const BIT_6 = &H1
Global Const BIT_7 = &H2
Global Const BIT_8 = &H3

Global Const STOP_1 = &H0                            ' Stop bits define
Global Const STOP_2 = &H4

Global Const P_EVEN = &H18                           ' Parity define
Global Const P_ODD = &H8
Global Const P_SPC = &H38
Global Const P_MRK = &H28
Global Const P_NONE = &H0

'       MODEM CONTROL setting
Global Const C_DTR = &H1
Global Const C_RTS = &H2

'       MODEM LINE STATUS
Global Const S_CTS = &H1
Global Const S_DSR = &H2
Global Const S_RI = &H4
Global Const S_CD = &H8

'  error code
Global Const SIO_OK = 0
Global Const SIO_BADPORT = -1        ' no such port or port not opened
Global Const SIO_OUTCONTROL = -2     ' can't control MOXA board
Global Const SIO_NODATA = -4         ' no data to read or no buffer to write
Global Const SIO_BADPARM = -7        ' bad parameter
Global Const SIO_WIN32FAIL = -8      ' call win32 function fail, please call
                                     ' GetLastError to get the error code

Declare Function sio_ioctl Lib "api232.dll" (ByVal port As Long, ByVal baud As Long, ByVal mode As Long) As Long
Declare Function sio_getch Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_read Lib "api232.dll" (ByVal port As Long, ByVal buf As String, ByVal Length As Long) As Long
Declare Function sio_putch Lib "api232.dll" (ByVal port As Long, ByVal term As Long) As Long
Declare Function sio_write Lib "api232.dll" (ByVal port As Long, ByVal buf As String, ByVal Length As Long) As Long
Declare Function sio_flush Lib "api232.dll" (ByVal port As Long, ByVal func As Long) As Long
Declare Function sio_iqueue Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_oqueue Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_lstatus Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_lctrl Lib "api232.dll" (ByVal port As Long, ByVal mode As Long) As Long
Declare Function sio_break Lib "api232.dll" (ByVal port As Long, ByVal time As Long) As Long
Declare Function sio_flowctrl Lib "api232.dll" (ByVal port As Long, ByVal mode As Long) As Long
Declare Function sio_Tx_hold Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_close Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_open Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_getbaud Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_getmode Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_getflow Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_DTR Lib "api232.dll" (ByVal port As Long, ByVal mode As Long) As Long
Declare Function sio_RTS Lib "api232.dll" (ByVal port As Long, ByVal mode As Long) As Long
Declare Function sio_baud Lib "api232.dll" (ByVal port As Long, ByVal Speed As Long) As Long
Declare Function sio_data_status Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_putb Lib "api232.dll" Alias "sio_write" (ByVal port As Long, ByVal buf As String, ByVal Length As Long) As Long
Declare Function sio_linput Lib "api232.dll" (ByVal port As Long, ByVal buf As String, ByVal Length As Long, ByVal term As Long) As Long
Declare Function ExecProgram Lib "vbc.dll" (ByVal hwndCaller As Long, ByVal pszCmdLine As String) As Long
Declare Function Fen_Song_Api Lib "vbc.dll" () As Long
Declare Function Vround Lib "vbc.dll" (ByVal x As Single, ByVal y As Long) As Single
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Declare Function can_read Lib "z_lj_dll.dll" (ByVal pei_number As String, ByVal Name As String, ByVal ban As Long, ByVal total_che As Long, ByVal Flag As Long) As Long
Declare Function read_power Lib "z_lj_dll.dll" (duan_hao As Integer, Power As Single, tempro As Integer, Flag As Integer) As Integer
Declare Function Set_Turn Lib "z_lj_dll.dll" (ByVal elec_name As Long, ByVal data As Long) As Long
Declare Function Set_Light Lib "z_lj_dll.dll" (ByVal elec_name As Long, ByVal data As Long) As Long


Type comm_get_buffer_type
    port As Integer
    tail As Integer
    head As Integer
    buffer(1024)  As Byte
End Type
Global comm_get_buffer  As comm_get_buffer_type

Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_ACTIVATE = &H6
Public Const WM_CREATE = &H1
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_GETFONT = &H31
Public Const WM_DESTROY = &H2
Public Const WM_PAINT = &HF

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_CHILD = 5


Declare Function GetTopWindow Lib "user32" (ByVal Hwnd As Long) As Long
Declare Function GetNextWindow Lib "user32" (ByVal Hwnd As Long, ByVal wFlag As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (lpClassName As Any, ByVal lpWindowName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)

'return true  is ok
' password check
Public Function Password_Do(modal_name As String, Prompt As String) As Boolean
    frmLogin.modal_name = modal_name
    frmLogin.Disp_Prompt = Prompt & frmLogin.Caption
'    Call SetWindowPos(frmLogin.Hwnd, HWND_TOPMOST)
    frmLogin.Show vbModal, MAIN_MDI

    If frmLogin.LoginSucceeded = True Then
        Password_Do = True
    Else
        Password_Do = False
    End If
End Function


