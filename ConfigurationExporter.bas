Attribute VB_Name = "ConfigurationExporter"
Option Explicit

'================================================================================================
'                         SOLIDWORKS DOCUMENT CONFIGURATION EXPORTER
'================================================================================================

' This module will extract solid models from each Solidworks part document and exported into separate files.
' Exported files will be saved as .Step.
' Exported files will be saved to the to the same directory in which the part file resides.
' Exported files will be saved as file name: [ConfigurationName].Step
'------------------------------------------------------------------------------------------------

'File extension used for the exported file name
Const ExportFileExt As String = ".step"

Sub Main()
    Dim ConfigName      As Variant
    For Each ConfigName In GetConfigurationNames
        SwDoc.ShowConfiguration2 (ConfigName)
        Call SwDoc.SaveAs(GetExportPath(CStr(ConfigName)))
    Next
    
    Call MsgBox("We are done here!", vbOKOnly + vbInformation, "Complete")
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

'Assembles the filename for the configuration export
Private Function GetExportPath(ConfigurationName As String)
    GetExportPath = GetDocDir & ConfigurationName & ExportFileExt
End Function
