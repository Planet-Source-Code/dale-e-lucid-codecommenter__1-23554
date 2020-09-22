VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   6600
   ClientLeft      =   1860
   ClientTop       =   540
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   11642
   _Version        =   393216
   Description     =   "Visual Basic v6.0 source code comment automation assistant will comment your code for easier coding and standardization."
   DisplayName     =   "Code Commenter"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
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

Const guidMYTOOL$ = "_C_O_D_E__C_O_M_E_N_T_E_R_"

Public FormDisplayed                As Boolean
Public VBInstance                   As VBIDE.VBE
Public WithEvents MenuHandler       As CommandBarEvents 'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents ToolbarHandler    As CommandBarEvents
Attribute ToolbarHandler.VB_VarHelpID = -1
Dim mcbMenuCommandBar               As Office.CommandBarControl
Dim mcbMenuCommandBar2              As Office.CommandBarControl
Dim mFCodeComenter                  As New FCodeComenter
'
'


Public Property Get NonModalApp() As Boolean
  NonModalApp = False  'used by addin toolbar
End Property


Sub Show()
  
    On Error Resume Next
    
    If mFCodeComenter Is Nothing Then
        Set mFCodeComenter = New FCodeComenter
    End If
    
    Set mFCodeComenter.VBInstance = VBInstance
    Set mFCodeComenter.Connect = Me
    FormDisplayed = True
    mFCodeComenter.Show vbModal
   
End Sub


'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo AddinInstance_OnConnection_Error
    
    'save the vb instance
    Set VBInstance = Application
    
    'menu(s) handler setup
    Call AddToCommandBar
  
    Exit Sub
    
AddinInstance_OnConnection_Error:
    MsgBox Err.Description
    
End Sub


'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    Dim cbMenu       As Object


    'delete the command bar entry
    mcbMenuCommandBar.Delete
    mcbMenuCommandBar2.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mFCodeComenter
    Set mFCodeComenter = Nothing

End Sub


Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    'set this to display the form on connect
End Sub


'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    'if this object isn't set then there's no code window to add code to
    If VBInstance.ActiveCodePane Is Nothing Then
        MsgBox "You must have a code window active to use this add-in.", vbInformation, "Code Comenter"
    Else
        Me.Show
    End If
End Sub


'this event fires when the toolbar is clicked in the IDE
Private Sub ToolbarHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call MenuHandler_Click(CommandBarControl, handled, CancelDefault)
End Sub


Private Sub AddToCommandBar()
Dim cbMenu              As Object

    On Error GoTo AddToCommandBar_Error

    'make sure the standard toolbar is visible
    VBInstance.CommandBars(2).Visible = True

    'add it to the command bar
    'the following line will add the Code Commenter to the
    'Standard toolbar to the right of the ToolBox button
    Set mcbMenuCommandBar = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
    'set the caption
    mcbMenuCommandBar.Caption = LoadResString(100)
    'copy the icon to the clipboard
    Clipboard.SetData LoadResPicture(9000, 0)
    'set the icon for the button
    mcbMenuCommandBar.PasteFace
    'sink the event
    Set Me.ToolbarHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    
    'see if we can find the Add-Ins menu so we can add the tool
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
        
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Sub
    End If
    
    'add it to the command bar
    Set mcbMenuCommandBar2 = cbMenu.Controls.Add(1)
    'set the caption
    mcbMenuCommandBar2.OnAction = "MenuHandler_Click(cbMenuCommandBar, True, True)"
    mcbMenuCommandBar2.Caption = LoadResString(100)
    mcbMenuCommandBar2.ShortcutText = "Ctrl+Shift+M"
    'sink the event
    Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar2)
    
    Exit Sub
    
AddToCommandBar_Error:
    MsgBox Err.Description
End Sub

