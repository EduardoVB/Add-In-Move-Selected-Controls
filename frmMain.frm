VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move controls"
   ClientHeight    =   2508
   ClientLeft      =   9996
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
   ScaleHeight     =   2508
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin ComCtl2.UpDown udnLeft 
      Height          =   360
      Left            =   2000
      TabIndex        =   10
      Top             =   552
      Width           =   252
      _ExtentX        =   445
      _ExtentY        =   635
      _Version        =   327681
      OrigLeft        =   2064
      OrigTop         =   552
      OrigRight       =   2316
      OrigBottom      =   912
      Max             =   0
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdUndoAll 
      Caption         =   "QQ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   200
      TabIndex        =   7
      ToolTipText     =   "Undo all moves made to the current selection"
      Top             =   1656
      Width           =   732
   End
   Begin VB.CommandButton cmdUndoLast 
      Caption         =   "Q"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   1032
      TabIndex        =   6
      ToolTipText     =   "Undo the last move"
      Top             =   1656
      Width           =   444
   End
   Begin VB.Timer tmrUpdatekSelection 
      Interval        =   300
      Left            =   72
      Top             =   792
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move"
      Enabled         =   0   'False
      Height          =   468
      Left            =   1566
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
   Begin ComCtl2.UpDown udnTop 
      Height          =   360
      Left            =   2000
      TabIndex        =   11
      Top             =   1008
      Width           =   252
      _ExtentX        =   445
      _ExtentY        =   635
      _Version        =   327681
      OrigLeft        =   2064
      OrigTop         =   552
      OrigRight       =   2316
      OrigBottom      =   912
      Max             =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lblMoved 
      AutoSize        =   -1  'True
      Caption         =   "0, 0"
      Height          =   240
      Left            =   1754
      TabIndex        =   9
      Top             =   2232
      Width           =   276
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Current selection moved:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   60
      TabIndex        =   8
      Top             =   2256
      Width           =   1596
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

Private mCurrentSelectionStr As String
Private mLastSelectionStr As String
Private mLeftMoves() As Single
Private mTopMoves() As Single
Private mScaleModeMoves As Long
Private mIncrement As Single

Private Sub cmdMove_Click()
    Dim iActualSelectionStr As String
    Dim iControlCount As Long
    
    iActualSelectionStr = GetCurrentSelectionStr(iControlCount)
    If iControlCount = 0 Then
        MsgBox "No control selected.", vbExclamation
        Exit Sub
    End If
    If iActualSelectionStr <> mLastSelectionStr Then
        ReDim mLeftMoves(0)
        ReDim mTopMoves(0)
        mLastSelectionStr = iActualSelectionStr
    End If
            
    ReDim Preserve mLeftMoves(UBound(mLeftMoves) + 1)
    mLeftMoves(UBound(mLeftMoves)) = CSng(txtLeft.Text)
    ReDim Preserve mTopMoves(UBound(mTopMoves) + 1)
    mTopMoves(UBound(mTopMoves)) = CSng(txtTop.Text)
    cmdUndoLast.Enabled = True
    cmdUndoAll.Enabled = True
    
    DoMove CSng(txtLeft.Text), CSng(txtTop.Text)
End Sub

Private Sub cmdUndoAll_Click()
    Dim iLeft As Single
    Dim iTop As Single
    
    CheckSelection
    If cmdUndoLast.Enabled = False Then
        MsgBox "Can't Undo. Restore the last selection."
        Exit Sub
    End If
    txtLeft.Text = mLeftMoves(1)
    txtTop.Text = mTopMoves(1)
    iLeft = -GetMovedLeft
    iTop = -GetMovedTop
    ReDim mLeftMoves(0)
    ReDim mTopMoves(0)
    DoMove iLeft, iTop
    CheckSelection
End Sub

Private Sub cmdUndoLast_Click()
    CheckSelection
    If cmdUndoLast.Enabled = False Then
        MsgBox "Can't Undo. Restore the last selection."
        Exit Sub
    End If
    txtLeft.Text = mLeftMoves(UBound(mLeftMoves))
    txtTop.Text = mTopMoves(UBound(mTopMoves))
    ReDim Preserve mLeftMoves(UBound(mLeftMoves) - 1)
    ReDim Preserve mTopMoves(UBound(mTopMoves) - 1)
    DoMove -CSng(txtLeft.Text), -CSng(txtTop.Text)
    CheckSelection
End Sub

Private Sub Form_Load()
    Me.Move GetSetting(App.Title, "Settings", "WindowLeft", Screen.Width * 0.5 - Me.Width), GetSetting(App.Title, "Settings", "WindowTop", Screen.Height * 0.4)
    ReDim mLeftMoves(0)
    ReDim mTopMoves(0)
    udnLeft.Increment = Screen.TwipsPerPixelX
    udnTop.Increment = Screen.TwipsPerPixelY
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "WindowLeft", Me.Left
    SaveSetting App.Title, "Settings", "WindowTop", Me.Top
    Connect.FormDisplayed = False
End Sub

Public Sub UpdateSelection()
    CheckSelection
End Sub

Private Sub CheckSelection()
    Dim c As Long
    
    mCurrentSelectionStr = GetCurrentSelectionStr(c)
    If (mCurrentSelectionStr = mLastSelectionStr) And (mLastSelectionStr <> "") And (UBound(mLeftMoves) > 0) Then
        cmdUndoLast.Enabled = True
        cmdUndoAll.Enabled = True
        lblMoved.Caption = GetMovedLeft & ", " & GetMovedTop
    Else
        cmdUndoLast.Enabled = False
        cmdUndoAll.Enabled = False
        lblMoved.Caption = "0, 0"
    End If
    If c > 0 Then
        lblTitle.Caption = "Move " & c & " control" & IIf(c = 1, "", "s") & ":"
        If IsNumeric(txtLeft.Text) And IsNumeric(txtTop.Text) Then
            cmdMove.Enabled = (CSng(txtLeft.Text) <> 0) Or (CSng(txtTop.Text) <> 0)
        Else
            cmdMove.Enabled = False
        End If
    Else
        lblTitle.Caption = "Move controls:"
        cmdMove.Enabled = False
    End If
End Sub
Private Function GetCurrentSelectionStr(Optional ControlCount As Long) As String
    Dim iCtl As VBControl
    Dim i As Single
    Dim iScaleModeComp As Integer
    
    ControlCount = 0
    On Error GoTo ErrH
    If Not VBInstance.ActiveVBProject Is Nothing Then
         GetCurrentSelectionStr = VBInstance.ActiveVBProject.Name & ","
         If Not VBInstance.SelectedVBComponent Is Nothing Then
            GetCurrentSelectionStr = GetCurrentSelectionStr & VBInstance.SelectedVBComponent.Name & ":"
            If VBInstance.SelectedVBComponent.HasOpenDesigner Then
                iScaleModeComp = VBInstance.SelectedVBComponent.Properties("ScaleMode").Value
                GetCurrentSelectionStr = GetCurrentSelectionStr & "(SM=" & iScaleModeComp & ")"
                mIncrement = Round(ScaleX(1, vbPixels, iScaleModeComp), 2)
                For Each iCtl In VBInstance.SelectedVBComponent.Designer.VBControls
                   If iCtl.InSelection Then
                       ControlCount = ControlCount + 1
                       GetCurrentSelectionStr = GetCurrentSelectionStr & iCtl.ControlObject.Name & "|"
                   End If
                Next
            End If
        End If
    End If
    Set iCtl = Nothing
    If ControlCount = 0 Then GetCurrentSelectionStr = ""
    Exit Function
    
ErrH:
    ControlCount = 0
    GetCurrentSelectionStr = ""
End Function

Private Sub tmrUpdatekSelection_Timer()
    CheckSelection
End Sub

Private Sub txtLeft_GotFocus()
    txtLeft.SelStart = 0
    txtLeft.SelLength = Len(txtLeft.Text)
End Sub

Private Sub txtLeft_Validate(Cancel As Boolean)
    If Not IsNumeric(txtLeft.Text) Then txtLeft.Text = "0"
End Sub

Private Sub txtTop_GotFocus()
    txtTop.SelStart = 0
    txtTop.SelLength = Len(txtTop.Text)
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
    Dim iDecSep As String
    
    iDecSep = Mid$(CStr(0.1), 2, 1)
    If InStr("-0123456789" & iDecSep & Chr(vbKeyBack), Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
    ElseIf (Chr$(KeyAscii) = "-") And (InStr(txtLeft.Text, "-") > 0) Then
        KeyAscii = 0
    ElseIf (Chr$(KeyAscii) = "-") And txtLeft.SelStart > 0 Then
        KeyAscii = 0
    ElseIf (Chr$(KeyAscii) = iDecSep) And (InStr(txtLeft.Text, iDecSep) > 0) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
    Dim iDecSep As String
    
    iDecSep = Mid$(CStr(0.1), 2, 1)
    If InStr("-0123456789" & iDecSep & Chr(vbKeyBack), Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
    ElseIf (Chr$(KeyAscii) = "-") And (InStr(txtTop.Text, "-") > 0) Then
        KeyAscii = 0
    ElseIf (Chr$(KeyAscii) = "-") And txtTop.SelStart > 0 Then
        KeyAscii = 0
    ElseIf (Chr$(KeyAscii) = iDecSep) And (InStr(txtTop.Text, iDecSep) > 0) Then
        KeyAscii = 0
    End If
End Sub

Private Sub DoMove(ByVal nLeft As Single, ByVal nTop As Single)
    Dim iCtl As VBControl
    
    If Not VBInstance.ActiveVBProject Is Nothing Then
         For Each iCtl In VBInstance.SelectedVBComponent.Designer.VBControls
            If iCtl.InSelection Then
                On Error Resume Next
                If TypeName(iCtl.ControlObject) = "Line" Then
                    If nLeft <> 0 Then
                        iCtl.ControlObject.X1 = iCtl.ControlObject.X1 + nLeft
                        iCtl.ControlObject.X2 = iCtl.ControlObject.X2 + nLeft
                    End If
                    If nTop <> 0 Then
                        iCtl.ControlObject.Y1 = iCtl.ControlObject.Y1 + nTop
                        iCtl.ControlObject.Y2 = iCtl.ControlObject.Y2 + nTop
                    End If
                Else
                    If nLeft <> 0 Then iCtl.ControlObject.Left = iCtl.ControlObject.Left + nLeft
                    If nTop <> 0 Then iCtl.ControlObject.Top = iCtl.ControlObject.Top + nTop
                End If
                On Error GoTo 0
            End If
         Next
    End If
    
    lblMoved.Caption = GetMovedLeft & ", " & GetMovedTop
End Sub

Private Function GetMovedLeft() As Single
    Dim c As Long
    
    For c = 1 To UBound(mLeftMoves)
        GetMovedLeft = GetMovedLeft + mLeftMoves(c)
    Next
End Function

Private Function GetMovedTop() As Single
    Dim c As Long
    
    For c = 1 To UBound(mTopMoves)
        GetMovedTop = GetMovedTop + mTopMoves(c)
    Next
End Function

Private Sub txtTop_Validate(Cancel As Boolean)
    If Not IsNumeric(txtTop.Text) Then txtTop.Text = "0"
End Sub

Private Sub udnLeft_DownClick()
    Dim v As Single
    
    v = CSng(txtLeft.Text) - mIncrement
    If Round(v) = v Then
        txtLeft.Text = v
    Else
        txtLeft.Text = Format(v, "0.0####")
    End If
End Sub

Private Sub udnLeft_UpClick()
    Dim v As Single
    
    v = CSng(txtLeft.Text) + mIncrement
    If Round(v) = v Then
        txtLeft.Text = v
    Else
        txtLeft.Text = Format(v, "0.0####")
    End If
End Sub

Private Sub udnTop_DownClick()
    Dim v As Single
    
    v = CSng(txtTop.Text) - mIncrement
    If Round(v) = v Then
        txtTop.Text = v
    Else
        txtTop.Text = Format(v, "0.0####")
    End If
End Sub

Private Sub udnTop_UpClick()
    Dim v As Single
    
    v = CSng(txtTop.Text) + mIncrement
    If Round(v) = v Then
        txtTop.Text = v
    Else
        txtTop.Text = Format(v, "0.0####")
    End If
End Sub

