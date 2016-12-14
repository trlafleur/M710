Attribute VB_Name = "MessageHook"
'
' Project..........Icom Control Panel
' File Name........MESSAGEHOOK.BAS
' File Version.....4/3/01
' Contents.........General purpose routine for hooking Windows messages...
'
' Copyright (c) 2001 - All Rights Reserved
' Victor Poor, W5SMM
' 1208 East River Drive, #302
' Melbourne, FL 32901
'
Option Explicit

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4
Public Const WM_MOUSEWHEEL = 522
Public bIsHooked As Boolean
Global lpPrevWndProc As Long
Global hMain As Long
Global lStep As Long

Public Sub Hook()
   If bIsHooked Then
      MsgBox "Don't hook it twice without unhooking, or you will be unable to unhook it."
   Else
      lpPrevWndProc = SetWindowLong(hMain, GWL_WNDPROC, AddressOf WindowProc)
      bIsHooked = True
   End If
End Sub

Public Sub Unhook()
   SetWindowLong hMain, GWL_WNDPROC, lpPrevWndProc
   bIsHooked = False
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' Place message intercept code here...
   Select Case uMsg
      Case WM_MOUSEWHEEL
         If wParam > 0 Then lStep = lStep - 1 Else lStep = lStep + 1
   End Select
   
   WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function


