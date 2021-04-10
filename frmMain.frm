VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move controls"
   ClientHeight    =   2328
   ClientLeft      =   9984
   ClientTop       =   2208
   ClientWidth     =   2700
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2328
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   2160
      Top             =   744
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   468
      Left            =   1536
      TabIndex        =   5
      Top             =   1656
      Width           =   900
   End
   Begin VB.TextBox txtTop 
      Height          =   360
      Left            =   1104
      TabIndex        =   4
      Text            =   "0"
      Top             =   1008
      Width           =   900
   End
   Begin VB.TextBox txtLeft 
      Height          =   360
      Left            =   1104
      TabIndex        =   2
      Text            =   "0"
      Top             =   552
      Width           =   900
   End
   Begin VB.Label lblTitle 
      Caption         =   "Move controls:"
      Height          =   324
      Left            =   168
      TabIndex        =   0
      Top             =   168
      Width           =   2220
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Top:"
      Height          =   324
      Left            =   480
      TabIndex        =   3
      Top             =   1056
      Width           =   468
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Left:"
      Height          =   324
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   468
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Sub cmdApply_Click()
    DoMove
End Sub

Private Sub Form_Load()
    Me.Move GetSetting(App.Title, "Settings", "WindowLeft", Screen.Width * 0.5 - Me.Width), GetSetting(App.Title, "Settings", "WindowTop", Screen.Height * 0.4)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "WindowLeft", Me.Left
    SaveSetting App.Title, "Settings", "WindowTop", Me.Top
End Sub

Private Sub Timer1_Timer()
    Dim iCtl As VBControl
    Dim c As Long
    
    lblTitle.Caption = "Move controls:"
    c = 0
    If Not VBInstance.ActiveVBProject Is Nothing Then
         For Each iCtl In VBInstance.SelectedVBComponent.Designer.VBControls
            If iCtl.InSelection Then
                c = c + 1
            End If
         Next
    End If
    If c > 0 Then
        lblTitle.Caption = "Move " & c & " control" & IIf(c = 1, "", "s") & ":"
        cmdApply.Enabled = True
    Else
        cmdApply.Enabled = False
    End If
End Sub

Private Sub txtLeft_GotFocus()
    txtLeft.SelStart = 0
    txtLeft.SelLength = Len(txtLeft.Text)
End Sub

Private Sub txtTop_GotFocus()
    txtTop.SelStart = 0
    txtTop.SelLength = Len(txtTop.Text)
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
    If InStr("-0123456789.", Chr$(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
    If InStr("-0123456789.", Chr$(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub DoMove()
    Dim iCtl As VBControl
    Dim iLeft As Single
    Dim iTop As Single
    
    iLeft = Val(txtLeft.Text)
    iTop = Val(txtTop.Text)
    
    lblTitle.Caption = "Move controls:"
    If Not VBInstance.ActiveVBProject Is Nothing Then
         For Each iCtl In VBInstance.SelectedVBComponent.Designer.VBControls
            If iCtl.InSelection Then
                On Error Resume Next
                If TypeName(iCtl.ControlObject) = "Line" Then
                    iCtl.ControlObject.X1 = iCtl.ControlObject.X1 + iLeft
                    iCtl.ControlObject.X2 = iCtl.ControlObject.X2 + iLeft
                    iCtl.ControlObject.Y1 = iCtl.ControlObject.Y1 + iTop
                    iCtl.ControlObject.Y2 = iCtl.ControlObject.Y2 + iTop
                Else
                    iCtl.ControlObject.Left = iCtl.ControlObject.Left + iLeft
                    iCtl.ControlObject.Top = iCtl.ControlObject.Top + iTop
                End If
                On Error GoTo 0
            End If
         Next
    End If
End Sub
