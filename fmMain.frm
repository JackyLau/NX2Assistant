VERSION 5.00
Begin VB.Form fmMain 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "NX 2 Assistant"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   1635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H0000FFFF&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1020
      TabIndex        =   1
      Top             =   220
      Width           =   550
   End
End
Attribute VB_Name = "fmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
 (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function SetWindowPos Lib "user32" _
 (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1

Const vTimeDelay As Integer = 150
Const vSendCommand As String = "^+S|~|~|~|%F|C"
Const vProgramName As String = "Capture NX 2"
Const vSaveName As String = "儲存選項"

Dim hWndFound As Long
Dim i As Byte
Dim n_time As Long

Dim c_AllText() As String
Dim c_Text As Variant



Private Sub cmdGo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        cmdGo.Enabled = False
        lblCount.Caption = ""
        
        hWndFound = FindWindow(vbNullString, vProgramName)
        
        If hWndFound <> 0 Then
            SetForegroundWindow hWndFound
            SendKeys "%F"
            Call p_Delay
            SendKeys "S"
            lblCount.Caption = "Raw"
        End If
        
        cmdGo.Enabled = True
    End If
End Sub


Private Sub cmdGo_Click()
    cmdGo.Enabled = False
    lblCount.Caption = ""
    
    hWndFound = FindWindow(vbNullString, vProgramName)
    
    If hWndFound <> 0 Then
        SetForegroundWindow hWndFound
'        Call p_VerOne
        Call p_VerTwo
    End If
    
    cmdGo.Enabled = True
End Sub


Private Sub p_VerTwo()
        
    SendKeys "^+S"
    Call p_Delay
    
    SendKeys "~"
    Call p_Delay
    
    If (FindWindow(vbNullString, "警告") <> 0) Or (FindWindow(vbNullString, "Warning") <> 0) Then
        SendKeys "~"
        Call p_Delay
    End If
    
    If (FindWindow(vbNullString, "儲存選項") <> 0) Or (FindWindow(vbNullString, "Save Option") <> 0) Then
        SendKeys "~"
        Call p_Delay
    Else
        Exit Sub
    End If
        
    For i = 5 To Int(99 / vTimeDelay * 100)
        If (FindWindow(vbNullString, "儲存選項") = 0) And (FindWindow(vbNullString, "Save Option") = 0) Then
            SetForegroundWindow hWndFound
            Exit For
        End If
        lblCount.Caption = Format((i * vTimeDelay / 1000), "0.0s")
        DoEvents
        Call p_Delay
    Next i
        
    SendKeys "%F"
    Call p_Delay
    SendKeys "C"
End Sub


Private Sub p_VerOne()
    For Each c_Text In c_AllText
        If c_Text = "%F" Then
            For i = UBound(c_AllText) To Int(99 / vTimeDelay * 100)
                If FindWindow(vbNullString, vSaveName) = 0 Then
                    SetForegroundWindow hWndFound
                    Exit For
                End If
                lblCount.Caption = Format((i * vTimeDelay / 1000), "0.0s")
                DoEvents
                Call p_Delay
            Next i
        End If
        Call p_Delay
        SendKeys c_Text
    Next
End Sub


Private Sub p_Delay()
    n_time = GetTickCount + vTimeDelay
    Do
        DoEvents
    Loop While GetTickCount() < n_time
End Sub


Private Sub Form_Load()
    Dim vCount As Integer
    
    If App.PrevInstance Then End
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    c_AllText = Split(vSendCommand, "|")
End Sub


