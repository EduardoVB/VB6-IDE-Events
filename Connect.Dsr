VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9948
   ClientLeft      =   1740
   ClientTop       =   1548
   ClientWidth     =   6588
   _ExtentX        =   11621
   _ExtentY        =   17547
   _Version        =   393216
   Description     =   "Displays a window that shows IDE events"
   DisplayName     =   "IDE events"
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

Private mAddInWindowVisible As Boolean
Private mAddInWindowLoaded As Boolean
Private mVBInstance As VBIDE.VBE
Private mcbMenuCommandBar As Office.CommandBarControl
Private mfrmAddInWindow As frmAddInWindow

Private WithEvents mMenuHandler As CommandBarEvents
Attribute mMenuHandler.VB_VarHelpID = -1

Private WithEvents mSelectedControls As VBIDE.SelectedVBControlsEvents
Attribute mSelectedControls.VB_VarHelpID = -1
Private WithEvents mControls As VBIDE.VBControlsEvents
Attribute mControls.VB_VarHelpID = -1
Private WithEvents mComponents As VBIDE.VBComponentsEvents
Attribute mComponents.VB_VarHelpID = -1
Private WithEvents mProjects As VBIDE.VBProjectsEvents
Attribute mProjects.VB_VarHelpID = -1
Private WithEvents mIDE As VBIDE.VBBuildEvents
Attribute mIDE.VB_VarHelpID = -1
Private WithEvents mFiles As VBIDE.FileControlEvents
Attribute mFiles.VB_VarHelpID = -1
Private WithEvents mReferences As VBIDE.ReferencesEvents
Attribute mReferences.VB_VarHelpID = -1

Public Sub HideWindow()
    On Error Resume Next
    
    mAddInWindowVisible = False
    mfrmAddInWindow.Hide
End Sub

Public Sub ShowWindow()
    If mAddInWindowVisible Then Exit Sub
    On Error Resume Next
    
    mAddInWindowVisible = True
    mfrmAddInWindow.Show
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    Set mVBInstance = Application
    
    If mfrmAddInWindow Is Nothing Then
        Set mfrmAddInWindow = New frmAddInWindow
        mAddInWindowLoaded = True
    End If
    
    Set mfrmAddInWindow.VBInstance = mVBInstance
    Set mfrmAddInWindow.Connect = Me
    
    If GetSetting(App.Title, "Settings", "ShowWindow", "0") = "1" Then
        ShowWindow
    End If
    
    LogEvent "IDE instance opened, EXE path: " & mVBInstance.FullName
    
    If ConnectMode <> ext_cm_External Then
        Set mcbMenuCommandBar = AddToAddInCommandBar("Show IDE event window")
        Set mMenuHandler = mVBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    Exit Sub
    
error_handler:
    MsgBox Err.Description
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    LogEvent "IDE instance closed, EXE path: " & mVBInstance.FullName
    
    mcbMenuCommandBar.Delete
    
    If mAddInWindowVisible Then
        SaveSetting App.Title, "Settings", "ShowWindow", "1"
        mAddInWindowVisible = False
    Else
        SaveSetting App.Title, "Settings", "ShowWindow", "0"
    End If
    
    Unload mfrmAddInWindow
    Set mfrmAddInWindow = Nothing
End Sub

Private Sub mComponents_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
    LogEvent "Added " & VBComponent.Name & " to " & VBComponent.VBE.ActiveVBProject.Name & " project"
End Sub

Private Sub mComponents_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
    LogEvent "Reloaded " & VBComponent.Name & " in " & VBComponent.VBE.ActiveVBProject.Name & " project"
End Sub

Private Sub mComponents_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    LogEvent "Removed " & VBComponent.Name & " from " & VBComponent.VBE.ActiveVBProject.Name & " project"
End Sub

Private Sub mComponents_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    LogEvent "Renamed " & OldName & " of project " & VBComponent.VBE.ActiveVBProject.Name & " to " & VBComponent.Name
End Sub

Private Sub mControls_ItemAdded(ByVal VBControl As VBIDE.VBControl)
    LogEvent "Added " & GetNameWithIndex(VBControl.ControlObject.Name, VBControl.ControlObject.Index) & " to " & VBControl.VBE.SelectedVBComponent.Name & " of " & VBControl.VBE.ActiveVBProject.Name & " project"
End Sub

Private Sub mControls_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
    LogEvent "Removed " & GetNameWithIndex(VBControl.ControlObject.Name, VBControl.ControlObject.Index) & " from " & VBControl.VBE.SelectedVBComponent.Name & " of " & VBControl.VBE.ActiveVBProject.Name & " project"
End Sub

Private Sub mControls_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
    LogEvent "Renamed " & GetNameWithIndex(OldName, OldIndex) & " to " & GetNameWithIndex(VBControl.ControlObject.Name, VBControl.ControlObject.Index) & " in " & VBControl.VBE.SelectedVBComponent.Name & " of " & VBControl.VBE.ActiveVBProject.Name & " project"
End Sub

Private Sub mFiles_AfterAddFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
    LogEvent "Added file " & FileName & " to " & VBProject & " project"
End Sub

Private Sub mFiles_AfterChangeFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal NewName As String, ByVal OldName As String)
    LogEvent "File named changed from " & OldName & " to " & NewName & " in project " & VBProject
End Sub

Private Sub mFiles_AfterCloseFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal WasDirty As Boolean)
    LogEvent "Closed file " & FileName & " in " & VBProject & " project"
End Sub

Private Sub mFiles_AfterRemoveFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
    LogEvent "Removed file " & FileName & " from " & VBProject & " project"
End Sub

Private Sub mFiles_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal Result As Integer)
    LogEvent "File " & FileName & " was written in " & VBProject & " project"
End Sub

Private Sub mFiles_BeforeLoadFile(ByVal VBProject As VBIDE.VBProject, FileNames() As String)
    Dim c As Long
    
    For c = 0 To UBound(FileNames)
        LogEvent "File " & FileNames(c) & " is about to be loaded in " & VBProject & " project"
    Next c
End Sub

Private Sub mFiles_DoGetNewFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, NewName As String, ByVal OldName As String, CancelDefault As Boolean)
    LogEvent "File named " & OldName & " is about to be changed to " & NewName & " in " & VBProject & " project"
End Sub

Private Sub mFiles_RequestChangeFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal NewName As String, ByVal OldName As String, Cancel As Boolean)
    LogEvent "File named " & OldName & " was requested to be changed to " & NewName & " in " & VBProject & " project"
End Sub

Private Sub mFiles_RequestWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileName As String, Cancel As Boolean)
    LogEvent "File " & FileName & " is requested to be written in " & VBProject & " project"
End Sub

Private Sub mReferences_ItemAdded(ByVal Reference As VBIDE.Reference)
    On Error Resume Next
    LogEvent "Reference " & Reference.Name & " with path: " & Reference.FullPath & " has been added to " & Reference.VBE.ActiveVBProject.Name & " project"
End Sub

Private Sub mReferences_ItemRemoved(ByVal Reference As VBIDE.Reference)
    On Error Resume Next
    LogEvent "Reference " & Reference.Name & " with path: " & Reference.FullPath & " has been removed from " & Reference.VBE.ActiveVBProject.Name & " project"
End Sub

Private Sub mSelectedControls_ItemAdded(ByVal VBControl As VBIDE.VBControl)
    LogEvent "Selected " & GetNameWithIndex(VBControl.ControlObject.Name, VBControl.ControlObject.Index) & " on " & VBControl.VBE.SelectedVBComponent.Name & " of " & VBControl.VBE.ActiveVBProject.Name & " project"
End Sub

Private Sub mSelectedControls_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
    LogEvent "Unselected " & GetNameWithIndex(VBControl.ControlObject.Name, VBControl.ControlObject.Index) & " on " & VBControl.VBE.SelectedVBComponent.Name & " of " & VBControl.VBE.ActiveVBProject.Name & " project"
End Sub

Private Sub mIDE_BeginCompile(ByVal VBProject As VBIDE.VBProject)
    LogEvent "IDE begins compiling " & mVBInstance.ActiveVBProject.Name & " project" ' VBProject can be Nothing, better use mVBInstance.ActiveVBProject instead here
End Sub

Private Sub mIDE_EnterDesignMode()
    LogEvent "IDE enters Design Mode"
End Sub

Private Sub mIDE_EnterRunMode()
    LogEvent "IDE enters Run Mode"
End Sub

Private Sub mMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ShowWindow
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    Set cbMenu = mVBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        Exit Function
    End If
    
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.Caption = sCaption
    Set AddToAddInCommandBar = cbMenuCommandBar
    Exit Function
    
AddToAddInCommandBarErr:
End Function

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    LogEvent "IDE instance startup complete, EXE path: " & mVBInstance.FullName
    SetMainhandlers
End Sub

Private Sub mComponents_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
    LogEvent "Activated " & VBComponent.Name & " of " & VBComponent.VBE.ActiveVBProject.Name & " project"
    Set mSelectedControls = mVBInstance.Events.SelectedVBControlsEvents(mVBInstance.ActiveVBProject, mVBInstance.SelectedVBComponent.Designer)
    Set mControls = mVBInstance.Events.VBControlsEvents(mVBInstance.ActiveVBProject, mVBInstance.SelectedVBComponent.Designer)
End Sub

Private Sub mComponents_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    LogEvent "Selected " & VBComponent.Name & " of " & VBComponent.VBE.ActiveVBProject.Name & " project"
    Set mSelectedControls = mVBInstance.Events.SelectedVBControlsEvents(mVBInstance.ActiveVBProject, mVBInstance.SelectedVBComponent.Designer)
    Set mControls = mVBInstance.Events.VBControlsEvents(mVBInstance.ActiveVBProject, mVBInstance.SelectedVBComponent.Designer)
End Sub

Private Sub mProjects_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    SetActiveObjectsHandlers
    LogEvent "Project " & VBProject.Name & " Activated"
End Sub

Private Sub mProjects_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    SetActiveObjectsHandlers
    LogEvent "Project " & VBProject.Name & " Added"
End Sub

Private Sub mProjects_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    SetActiveObjectsHandlers
    LogEvent "Project " & VBProject.Name & " Removed"
End Sub

Private Sub mProjects_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    SetActiveObjectsHandlers
    LogEvent "Project " & OldName & " Renamed to " & VBProject.Name
End Sub

Private Sub SetActiveObjectsHandlers()
    If Not mVBInstance.ActiveVBProject Is Nothing Then
        Set mComponents = Nothing
        Set mComponents = mVBInstance.Events.VBComponentsEvents(mVBInstance.ActiveVBProject)
    End If
    
    If Not mVBInstance.SelectedVBComponent Is Nothing Then
        If mVBInstance.SelectedVBComponent.HasOpenDesigner Then
            Set mSelectedControls = Nothing
            Set mSelectedControls = mVBInstance.Events.SelectedVBControlsEvents(mVBInstance.ActiveVBProject, mVBInstance.SelectedVBComponent.Designer)
            Set mControls = Nothing
            Set mControls = mVBInstance.Events.VBControlsEvents(mVBInstance.ActiveVBProject, mVBInstance.SelectedVBComponent.Designer)
        Else
            'No designer selected
        End If
    Else
        'No designer selected
    End If

    Set mFiles = Nothing
    Set mFiles = mVBInstance.Events.FileControlEvents(mVBInstance.ActiveVBProject)
    Set mReferences = Nothing
    Set mReferences = mVBInstance.Events.ReferencesEvents(mVBInstance.ActiveVBProject)
    
End Sub

Public Sub SetMainhandlers()
    Dim iEvents2 As Events2
    
    Set mProjects = Nothing
    Set mProjects = mVBInstance.Events.VBProjectsEvents
    Set mIDE = Nothing
    Set iEvents2 = mVBInstance.Events
    Set mIDE = iEvents2.VBBuildEvents
    
    SetActiveObjectsHandlers
End Sub

Private Sub LogEvent(nEventDescription As String)
    Debug.Print nEventDescription
    
    If mAddInWindowLoaded Then
        mfrmAddInWindow.txtActions.SelStart = Len(mfrmAddInWindow.txtActions.Text)
        mfrmAddInWindow.txtActions.SelText = nEventDescription & vbCrLf
        mfrmAddInWindow.ZOrder
    End If
End Sub

Private Function GetNameWithIndex(nName As String, nIndex As Long) As String
    GetNameWithIndex = nName
    If nIndex > -1 Then
        GetNameWithIndex = GetNameWithIndex & "(" & nIndex & ")"
    End If
End Function
