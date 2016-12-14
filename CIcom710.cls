VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIcom710"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Project..........Icom Control Panel
' File Name........CICOM710.CLS
' File Version.....4/3/01
' Contents.........Interface class for controlling an ICOM 710...
'
' Copyright (c) 2001 - All Rights Reserved
' Victor Poor, W5SMM
' 1208 East River Drive, #302
' Melbourne, FL 32901
'
' PTC II control's by Tom Lafleur, KA6IQA
'
Option Explicit

Public Enum Emission
   rmNBDP
   rmUSB
   rmLSB
   rmCW
   rmAM
End Enum

Private sRadio As String
Private sBaud As String
Private bPTC As Boolean
Private PTCcmd As Boolean
Private sp As MSComm

Private Function RadioCommand(sCommand As String) As Boolean
   Dim sBuffer As String
   Dim sHex As String
   Dim sCheckSum As String
   Dim lCheckSum As Long
   Dim lIndex As Long
   Dim lSeconds As Long
  
   RadioCommand = False
       If Not sp.PortOpen Then Exit Function
       
    If PTCcmd Then          ' Check if we need raw output to PTC II
        sBuffer = sCommand
        sp.Output = sCommand
    Else
       
        lCheckSum = 0
        sBuffer = "$PICOA,90," & sRadio & "," & sCommand
    
        For lIndex = 2 To Len(sBuffer)
            lCheckSum = lCheckSum Xor Asc(Mid(sBuffer, lIndex, 1))
        Next lIndex
        sCheckSum = "*" & Right("0" & Hex$(lCheckSum), 2)
    
    If bPTC Then                ' If PTC II connected
        sBuffer = sBuffer & sCheckSum & vbCrLf
        ' We need to convert Radio Commands to a HEX sequence
        ' sRet will hold the Command string in HEX
        sHex = ""
        For lIndex = 1 To Len(sBuffer)
            sHex = sHex & Right("0" & Hex(Asc(Mid(sBuffer, lIndex, 1))), 2)
        Next lIndex
        ' Commands sequence needed to pass via the PTC II radio port
        ' PTC Format is set dump mode, then special header, then command in hex.
        '      A max of 256 bytes can be sent. (512 hex char)
        ' Special header is: (CTL5) #TX:  hex data cr
        '      PTC return a max of 40 bytes in this mode
        sBuffer = Chr(5) & "#TX:" & sHex & vbCr
        sp.Output = sBuffer     ' Output Header and Radio Command String
    
    Else
    ' If connected via NMEA port
        sBuffer = sBuffer & sCheckSum & vbCrLf
        sp.Output = sBuffer
    End If
  End If
  
   lSeconds = (Second(Time) + 2) Mod 60
   Do
    If sp.InBufferCount >= Len(sBuffer) Then
        RadioCommand = True
        Exit Do
    End If
    DoEvents
    If Second(Time) = lSeconds Then
        Exit Function
    End If
   Loop
        sBuffer = sp.Input              ' Empty Input buffer
        
End Function

Public Sub OpenPort(spPort As MSComm, lPort As Long)
                  
   Set sp = spPort
   On Error Resume Next
         
   sp.PortOpen = False
   sp.CommPort = lPort + 1
   
   If bPTC Then
    ' Will need to change baud to sBaud if bPTC is true
    ' May need to hunt for correct baud rate on PTC II over a range
   
        sp.Settings = sBaud & ",N,8,1"
        sp.PortOpen = True                      ' Open the I/O Port
        
     ' We want to wake up modem from it Autobaud mode, Set the Radio Baud rate
     ' We may want to look for a "CMD" return from PTC II to see ifts awake
        PTCcmd = True       ' Tell output routine that we want raw output to PTC
        RadioCommand vbCr                        ' Wake up Modem with a CR
        RadioCommand "TRX TY Icom 4800 3" & vbCr ' Tell PTC II we are talking to an ICOM at 4800 baud
        RadioCommand "TRX DU 1" & vbCr           ' Enable DUMP mode in PTC II
        PTCcmd = False
        
    Else
        sp.Settings = "4800,N,8,1"          ' NMEA standard for ICOM Direct connection
        sp.PortOpen = True
        
   End If

End Sub

Public Function OpenRadio() As Boolean
   If RadioCommand("REMOTE,ON") = False Then
      OpenRadio = False
      Exit Function
   Else
      OpenRadio = True
   End If
   RadioCommand "AGC,ON"
   RadioCommand "TXP,3"
   RadioCommand "DIM,OFF"
   
End Function

Public Function CloseRadio() As Boolean
   If RadioCommand("AFG,0") Then
      RadioCommand "RFG,9"
      RadioCommand "REMOTE,OFF"
      CloseRadio = True
   Else
      CloseRadio = False
   End If
   On Error Resume Next
   sp.PortOpen = False
   
End Function

Public Sub Mode(emMode As Emission)
   Select Case emMode
      Case rmNBDP
         RadioCommand "MODE,AFS"
      Case rmUSB
         RadioCommand "MODE,USB"
      Case rmLSB
         RadioCommand "MODE,LSB"
      Case rmCW
         RadioCommand "MODE,CW"
      Case rmAM
         RadioCommand "MODE,AM"
   End Select
   
End Sub

Public Sub AudioGain(lLevel As Long)
   RadioCommand "AFG," & CStr(lLevel)
   
End Sub

Public Sub RFGain(lLevel As Long)
   RadioCommand "RFG," & CStr(Int(lLevel / 11))
   
End Sub

Public Sub SetReceiver(dQRG As Double)
   RadioCommand "RXF," & Format(dQRG / 1000#, "#0.000000")
   
End Sub

Public Sub SetTransmitter(dQRG As Double)
   RadioCommand "TXF," & Format(dQRG / 1000#, "#0.000000")
   
End Sub

Public Sub PanelLight(bDim As Boolean)
   If bDim Then
      RadioCommand "DIM,ON"
   Else
      RadioCommand "DIM,OFF"
   End If
   
End Sub

Public Sub NoiseBlank(bNB As Boolean)
   If bNB Then
      RadioCommand "NB,ON"
   Else
      RadioCommand "NB,OFF"
   End If
   
End Sub

Public Sub Squelch(bSquelch As Boolean)
   If bSquelch Then
      RadioCommand "SQLC,ON"
   Else
      RadioCommand "SQLC,OFF"
   End If
   
End Sub

Property Let RadioType(sRadioType As String)
   sRadio = sRadioType
   
End Property
    
Property Let PTCBaud(sPTCbaud As String)
   sBaud = sPTCbaud
   
End Property

Property Let PTC(bPTCflag As Boolean)
    bPTC = bPTCflag
    
End Property


