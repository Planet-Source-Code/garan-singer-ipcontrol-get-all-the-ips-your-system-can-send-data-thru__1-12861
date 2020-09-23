VERSION 5.00
Begin VB.UserControl mcalpGetAllIPs 
   BackColor       =   &H80000004&
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   735
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   735
   ScaleWidth      =   735
   ToolboxBitmap   =   "mcalpGetAllIPs.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "mcalpGetAllIPs.ctx":0312
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "mcalpGetAllIPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim DidInit As Boolean

Private Sub UserControl_Terminate()
    If DidInit Then
        SocketsCleanup
    End If
End Sub

Private Sub SocketsCleanup()
    Dim lReturn As Long
    
    lReturn = WSACleanup()
    
    If lReturn <> 0 Then
        MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetIPs(ReturnedIPs() As String) As Integer
    Dim NumIPs As Integer
    
    NumIPs = 0
    
    If DidInit = False Then
        SocketsInit
        DidInit = True
    End If
    
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    
    If gethostname(hostname, 256) = SOCKET_ERROR Then
        MsgBox "Windows Sockets error " & Str(WSAGetLastError())
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If
    
    hostent_addr = gethostbyname(hostname)
    
    If hostent_addr = 0 Then
        MsgBox "Winsock.dll is not responding."
        Exit Function
    End If
    
    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4
    
    'MsgBox hostname
    
    'get all of the IP address if machine is  multi-homed
    
    Do
        NumIPs = NumIPs + 1
        ReDim temp_ip_address(1 To host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
        
        For i = 1 To host.hLength
            ip_address = ip_address & temp_ip_address(i) & "."
        Next
        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
        
        ReDim Preserve ReturnedIPs(NumIPs)
        ReturnedIPs(NumIPs) = ip_address
        
        ip_address = ""
        host.hAddrList = host.hAddrList + LenB(host.hAddrList)
        RtlMoveMemory hostip_addr, host.hAddrList, 4
    Loop While (hostip_addr <> 0)
    
    GetIPs = NumIPs
End Function

Private Sub SocketsInit()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String

    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

    If iReturn <> 0 Then
        MsgBox "Winsock.dll is not responding."
        Exit Sub
    End If

    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
             WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        MsgBox sMsg
        Exit Sub
    End If

    'iMaxSockets is not used in winsock 2. So the following check is only
    'necessary for winsock 1. If winsock 2 is requested,
    'the following check can be skipped.

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg
        Exit Sub
    End If

End Sub

Private Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

