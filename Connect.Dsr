VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9948
   ClientLeft      =   2412
   ClientTop       =   1212
   ClientWidth     =   6588
   _ExtentX        =   11621
   _ExtentY        =   17547
   _Version        =   393216
   Description     =   "Move the currently selected controls, all group togheter"
   DisplayName     =   "Move  controls"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Private mcbMenuCommandBar         As Office.CommandBarControl
Private mfrmMain                 As New frmMain
Public WithEvents MenuHandler As CommandBarEvents          'controlador de evento de barra de comandos
Attribute MenuHandler.VB_VarHelpID = -1
Private WithEvents mSelectedControls As VBIDE.SelectedVBControlsEvents
Attribute mSelectedControls.VB_VarHelpID = -1
Private WithEvents mControls As VBIDE.VBControlsEvents
Attribute mControls.VB_VarHelpID = -1
Private WithEvents mComponents As VBIDE.VBComponentsEvents
Attribute mComponents.VB_VarHelpID = -1
Private WithEvents mProjects As VBIDE.VBProjectsEvents
Attribute mProjects.VB_VarHelpID = -1

Sub HidefrmMain()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmMain.Hide
   
End Sub

Sub ShowfrmMain()
  
    On Error Resume Next
    
    If mfrmMain Is Nothing Then
        Set mfrmMain = New frmMain
    End If
    
    Set mfrmMain.VBInstance = VBInstance
    Set mfrmMain.Connect = Me
    FormDisplayed = True
    mfrmMain.Show
    mfrmMain.UpdateSelection
    mfrmMain.ZOrder
    mfrmMain.SetFocus
   
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    Set VBInstance = Application
    
    If ConnectMode = ext_cm_External Then
        ShowfrmMain
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar(App.Title)
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    mcbMenuCommandBar.Delete
    
    Unload mfrmMain
    Set mfrmMain = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        ShowfrmMain
    End If
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    SetMainhandlers
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ShowfrmMain
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'objeto de barra de comandos
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        Exit Function
    End If
    
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

Public Sub SetMainhandlers()
    Dim iEvents2 As Events2
    
    Set mProjects = Nothing
    Set mProjects = VBInstance.Events.VBProjectsEvents
    
    SetActiveObjectsHandlers
End Sub

Private Sub mProjects_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    SetActiveObjectsHandlers
End Sub

Private Sub mProjects_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    SetActiveObjectsHandlers
End Sub

Private Sub mProjects_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    SetActiveObjectsHandlers
End Sub

Private Sub mProjects_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    SetActiveObjectsHandlers
End Sub

Private Sub SetActiveObjectsHandlers()
    If Not VBInstance.ActiveVBProject Is Nothing Then
        Set mComponents = Nothing
        Set mComponents = VBInstance.Events.VBComponentsEvents(VBInstance.ActiveVBProject)
    End If
    
    If Not VBInstance.SelectedVBComponent Is Nothing Then
        If VBInstance.SelectedVBComponent.HasOpenDesigner Then
            Set mSelectedControls = Nothing
            Set mSelectedControls = VBInstance.Events.SelectedVBControlsEvents(VBInstance.ActiveVBProject, VBInstance.SelectedVBComponent.Designer)
            Set mControls = Nothing
            Set mControls = VBInstance.Events.VBControlsEvents(VBInstance.ActiveVBProject, VBInstance.SelectedVBComponent.Designer)
        Else
            'No designer selected
        End If
    Else
        'No designer selected
    End If
End Sub

Private Sub mSelectedControls_ItemAdded(ByVal VBControl As VBIDE.VBControl)
    If FormDisplayed Then
        mfrmMain.UpdateSelection
    End If
End Sub

Private Sub mSelectedControls_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
    If FormDisplayed Then
        mfrmMain.UpdateSelection
    End If
End Sub
