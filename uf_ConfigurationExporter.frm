VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_ConfigurationExporter 
   Caption         =   "Solidworks Configuration Exporter"
   ClientHeight    =   1440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7110
   OleObjectBlob   =   "uf_ConfigurationExporter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_ConfigurationExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================================
'                         SOLIDWORKS DOCUMENT CONFIGURATION EXPORTER
'================================================================================================

' This module will extract solid models from each Solidworks part document and exported into separate files.
' Exported files may be saved to the to the same directory in which the part file resides or into sub directories under the document directory.
' Exported files may be saved as file name: [ConfigurationName].Step or [ConfigurationName].x_b
'------------------------------------------------------------------------------------------------

'Creates the specified folder within a directory.
'If folder is preexisting, folder's contents will be deleted.
Private Sub CreateFolder(RootDir As String, FolderName As String)
    Dim Directory           As String:      Directory = RootDir & FolderName
    If Dir(Directory, vbDirectory) = "" Then
        Call MkDir(Directory)
    Else
        Kill Directory & "\*"
    End If
End Sub


Private Sub cb_Export_Click()

    'Validate User's inputs
    If Not ValidateForm Then
        Call MsgBox("Select at least one configuration and the export file type" & vbCrLf & "Please try again", vbExclamation + vbOKOnly, "Input Error")
        Exit Sub
    End If
    
    Dim UserConfiguration   As String:      UserConfiguration = SwDoc.ConfigurationManager.ActiveConfiguration.Name     '<== Currently active configuration
    Dim Configurations()    As String:      Configurations = GetConfigurationNames                                      '<== All document configurations
    Dim Configuration       As String
    Dim ExportFileName      As String
    Dim RootDir             As String:      RootDir = GetDocDir()                                                       '<== Root directory of the current document
    Dim i                   As Integer

    For i = 0 To UBound(Configurations)
        Configuration = Configurations(i)
        
        'Activate configuration
        Call SwDoc.ShowConfiguration2(Configuration)
        
        'Determine filename of the export
        If Me.cb_CreateFolders Then
            'Create subdirectory if needed
            Call CreateFolder(RootDir, Configuration)
            ExportFileName = RootDir & Configuration & "\" & Configuration & GetFileExt()
        Else
            ExportFileName = RootDir & Configuration & GetFileExt()
        End If

        'Export
        Call SwDoc.SaveAs(ExportFileName)
    Next i
    
    'Reinstate user's initial configuration
    Call SwDoc.ShowConfiguration2(UserConfiguration)
    
    Call MsgBox("We are done here!", vbOKOnly + vbInformation, "Complete")
End Sub

'Retrieves the file extension in accordance with the user's selection
Private Function GetFileExt() As String
    If Me.ob_Step Then
        GetFileExt = ".step"
    ElseIf Me.ob_Xt Then
        GetFileExt = ".x_b"
        End If
End Function

'Tests whether all necessary controls have been populated by the user
Private Function ValidateForm() As Boolean
    Dim ChkBox As CheckBox
    'Ensure at least one configuration checkbox has been ticked
    For Each ChkBox In fr_Configurations.Controls
        If ChkBox.value Then ValidateForm = True
    Next ChkBox
    'Ensure a export type has been selected
    If Not (ob_Step.value Or ob_Xt.value) Then ValidateForm = False
End Function

Private Sub UserForm_Initialize()
    PopulateConfigurations
End Sub

'Creates a checkbox control for each configuration with the solidworks document
Private Sub PopulateConfigurations()
    Dim i                   As Integer
    Dim Configurations()    As String:      Configurations = GetConfigurationNames
    
    For i = 0 To UBound(Configurations)
        Call CreateConfigurationCheckBox(Configurations(i), i)
    Next i
    
    ResizeForm (Configurations(i - 1))
End Sub

'Resizes form once all dynamic controls are generated
Private Sub ResizeForm(Configuration As String)
    Dim Height              As Integer:
    Height = Controls("cb_" & Configuration).Top + Controls("cb_" & Configuration).Height + 10
    If fr_Configurations.Height < Height Then fr_Configurations.Height = Height
    Me.Height = fr_Configurations.Top + fr_Configurations.Height + 35
End Sub

'Creates and positions a checkbox to allow the user to confirm configuration export
Private Sub CreateConfigurationCheckBox(ConfigurationName As String, Index As Integer)
    With Me.fr_Configurations.Controls.Add("Forms.CheckBox.1", "cb_" & ConfigurationName, False)
        .Caption = ConfigurationName
        .Left = 10
        .Top = Index * 18 + 10
        .Visible = True
        .value = True
    End With
End Sub

'Retrives the directory in which the active solidworks document resides.
'Trailing slash will be omitted if 'False' is passed as an argument.
Private Function GetDocDir(Optional IncludeSlash As Boolean = True) As String
    Dim DocFileName     As String:      DocFileName = SwDoc.GetPathName
    Dim SlashLocation   As Integer:     SlashLocation = InStrRev(DocFileName, "\")
    GetDocDir = Mid(DocFileName, 1, SlashLocation + CInt(Not IncludeSlash))
End Function

'Retrieves the solidworks application
Private Function SwApp() As SldWorks.SldWorks
    Set SwApp = SolidWorks.SldWorks
End Function

'Retrieves the active solidworks document
Private Function SwDoc() As SldWorks.ModelDoc2
    Set SwDoc = SwApp.ActiveDoc
End Function

'Retrieves all configuration names within a solidwork document
Private Function GetConfigurationNames() As String()
    GetConfigurationNames = SwDoc.GetConfigurationNames
End Function
