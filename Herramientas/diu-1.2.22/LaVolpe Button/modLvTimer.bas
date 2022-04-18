Attribute VB_Name = "modLvTimer"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

'***************************************************************************
'*  Name         : lv_TimerCallBack
'*  Description  :
'*  Parameters   : hWnd As Long, Message As Long,
'*               : wParam As Long, lParam As Long
'*  Returns      : Long
'*  Author       : GS
'*  Date         : 30 may 2010
'***************************************************************************

Public Function lv_TimerCallBack(ByVal hWnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'*************************************************
'Author: LaVolpe
'Last modified: 20/05/06
'*************************************************

Dim tgtButton As lvButtons_H
' when timer was intialized, the button control's hWnd
' had property set to the handle of the control itself
' and the timer ID was also set as a window property
CopyMemory tgtButton, GetProp(hWnd, "lv_ClassID"), &H4
Call tgtButton.TimerUpdate(GetProp(hWnd, "lv_TimerID"))  ' fire the button's event
CopyMemory tgtButton, 0&, &H4                                    ' erase this instance
End Function

