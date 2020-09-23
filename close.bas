Attribute VB_Name = "modCloseBtn"
Option Explicit

Private Const SC_CLOSE As Long = &HF060&
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const SC_MINIMIZE As Long = &HF020&

Private Const xSC_CLOSE As Long = -10&
Private Const xSC_MAXIMIZE As Long = -11&
Private Const xSC_MINIMIZE As Long = -12&

Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000

Private Const hWnd_NOTOPMOST = -2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_FRAMECHANGED = &H20

Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemInfo Lib "user32" Alias _
    "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias _
    "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Private Declare Function IsWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias _
    "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    
Private Declare Function SetWindowLong Lib "user32" Alias _
    "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) _
    As Long
    
Private Declare Function SetParent Lib "user32" ( _
    ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    
Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub EnableMaxButton(ByVal hWnd As Long, Enable As Boolean)

    
    Dim lngFormStyle As Long
    lngFormStyle = GetWindowLong(hWnd, GWL_STYLE)
    If Enable Then
        lngFormStyle = lngFormStyle Or WS_MAXIMIZEBOX
    Else
        lngFormStyle = lngFormStyle And Not WS_MAXIMIZEBOX
    End If
    SetWindowLong hWnd, GWL_STYLE, lngFormStyle
            
    SetParent hWnd, GetParent(hWnd)
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, _
            SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
End Sub

