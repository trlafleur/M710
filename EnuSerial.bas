Attribute VB_Name = "EnuSerial"
Option Explicit

' DCB to check for installed comm ports
'
Public Type DCB
   DCBlength As Long
   BaudRate As Long
   fBitFields As Long
   wReserved As Integer
   XonLim As Integer
   XoffLim As Integer
   ByteSize As Byte
   Parity As Byte
   StopBits As Byte
   XonChar As Byte
   XoffChar As Byte
   ErrorChar As Byte
   EofChar As Byte
   EvtChar As Byte
   wReserved1 As Integer
End Type

Public Type COMMCONFIG
   dwSize As Long
   wVersion As Integer
   wReserved As Integer
   dcbx As DCB
   dwProviderSubType As Long
   dwProviderOffset As Long
   dwProviderSize As Long
   wcProviderData As Byte
End Type

' Required to find all serial ports on a system
Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" _
    (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long

