VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelUiGuard"
Option Explicit

Private m_prevScreenUpdating As Boolean
Private m_prevEnableEvents As Boolean
Private m_prevStatusBar As Boolean
Private m_statusBarManaged As Boolean
Private m_restored As Boolean

Private Sub Class_Initialize()
    m_prevScreenUpdating = Application.ScreenUpdating
    m_prevEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

Public Sub DisableStatusBar()
    If m_statusBarManaged Then Exit Sub

    m_prevStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = False
    m_statusBarManaged = True
End Sub

Public Sub Restore()
    If m_restored Then Exit Sub

    Application.ScreenUpdating = m_prevScreenUpdating
    Application.EnableEvents = m_prevEnableEvents

    If m_statusBarManaged Then
        Application.DisplayStatusBar = m_prevStatusBar
    End If

    m_restored = True
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Restore
End Sub
