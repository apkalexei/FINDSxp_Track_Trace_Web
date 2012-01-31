VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReportExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' **** NOTE:  THIS ASSUMES THAT THE PARAMETERS ARE ALL LONG VALUES!!! ****
Option Explicit

Public Enum enmReportExportFileTypes
    FT_EXPORT_PDF = 0
    FT_EXPORT_WORD = 1
    FT_EXPORT_HTML = 2
End Enum

Private aryRptParamArray() As Variant
Private intCurrentParameter As Integer

Public Sub addParameter(str_in_Value As Variant)
    ' Pass the value as a string - the parameter number will automatically be created
    ' **** PARAMETERS MUST BE ADDED IN ORDER AND PASSED IN AS THE TYPE THAT IS EXPECTED BY CRYSTAL REPORTS ****
    intCurrentParameter = intCurrentParameter + 1
    ReDim Preserve aryRptParamArray(intCurrentParameter) As Variant
    aryRptParamArray(intCurrentParameter - 1) = str_in_Value
End Sub

Public Function getParameterArray() As String
    ' For debugging purposes, view the contents of the array in HTML format
    Dim i As Integer
    Dim strTemp As String
    strTemp = "Parameters: <BR>"
    
    For i = 1 To intCurrentParameter
        strTemp = strTemp & "#: " & i & "   Value: " & aryRptParamArray(i - 1) & "<br>"
    Next
    
    getParameterArray = strTemp
End Function

Public Sub ExportReport(ByVal str_in_ReportType As String, ByVal str_in_ReportFileName As String, ByVal str_in_OutFilename As String)
' We are checking for the following problems in error handling:
'   - Report file not found
'   - Exported document not found after exporting
'   - The number of parameters passed does not match the number of parameters required for the report
    
    On Error GoTo err_ExportReport
    ' Report Type: See Enumeration
    ' Report Filename: Fully qualified (c:\reports\bookconf.rpt)
    ' Output File Filename: Fully qualified (c:\documents\myreport.pdf)
    Dim i As Integer
    Dim Report As New CRAXDRT.Report
    Dim application As New CRAXDRT.application
    
    ' Used for counting parameters and creating an error based on the parameters
    Dim crpParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crpParamDef As CRAXDRT.ParameterFieldDefinition
    Dim intParamCount

    Dim fso, MyFile
    
    ' Check to make sure that the report file exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(str_in_ReportFileName) Then
        Err.Raise -12987437, , "Failed to fetch the report, '" & str_in_ReportFileName & "'.  The report file was not found."
    End If
    
    Set fso = Nothing
    
    ' Open the report
    Set Report = application.OpenReport(str_in_ReportFileName)
    ' Disable any messages - the web server can't deal with parameters being handled in this way
    Report.EnableParameterPrompting = False
    Report.MorePrintEngineErrorMessages = False
    
    ' Count the parameters in the actual report
    Set crpParamDefs = Report.ParameterFields
    For Each crpParamDef In crpParamDefs
        intParamCount = intParamCount + 1
    Next
    
    ' If it's not the same as the num params given to us, then raise an error.
    If intParamCount <> UBound(aryRptParamArray) Then
        Dim strParameterList
        ' Build a string of the parameters(For debugging)
        strParameterList = "<Parameter List>"
        For i = 1 To intCurrentParameter
            strParameterList = strParameterList & "'" & aryRptParamArray(i - 1) & "' "
        Next
        strParameterList = strParameterList & "</Parameter List>"
        Err.Raise -12345689, "clsReportExporter.ExportReport", "Wrong number of parameters passed to the report, " & str_in_ReportFileName & ".  Number of parameters in rpt: " & intParamCount & ", Num parameters passed to object: " & UBound(aryRptParamArray) & ".  " & strParameterList
    End If
    
    'Create the parameters in the report itself
    For i = 1 To intCurrentParameter
        Call Report.ParameterFields(i).SetCurrentValue(aryRptParamArray(i - 1))
    Next

    'Based on the format type (see enumeration), set variables in the report
    If str_in_ReportType = FT_EXPORT_PDF & "" Then
        Report.ExportOptions.FormatType = crEFTPortableDocFormat
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.DiskFileName = str_in_OutFilename
    ElseIf str_in_ReportType = FT_EXPORT_WORD & "" Then
        Report.ExportOptions.FormatType = crEFTWordForWindows
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.DiskFileName = str_in_OutFilename
    ElseIf str_in_ReportType = FT_EXPORT_HTML & "" Then
        Report.ExportOptions.FormatType = 32
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.HTMLFileName = str_in_OutFilename
    End If
    
    ' Export it!
    Report.Export False
    
    ' Check to make sure that the exported report file exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(str_in_OutFilename) Then
        Err.Raise -12987437, , "Failed to generate the report, '" & str_in_OutFilename & "'.  The exported report file was not found."
    End If
    
    Set fso = Nothing
    
    Exit Sub
err_ExportReport:
    Err.Raise -12345123, "clsReportExporter.ExportReport", "Error: " & Err.Description
End Sub

Private Sub Class_Initialize()
    intCurrentParameter = 0
End Sub
