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
Public cServer As String
Public cDataBase As String
Public cUser As String
Public cPWD As String
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
    Dim crxTables As CRAXDRT.DatabaseTables
    Dim crxTable As CRAXDRT.DatabaseTable
    Dim crxSubreportObject As CRAXDRT.SubreportObject
    Dim crxSubReport As CRAXDRT.Report
    Dim crxSections As CRAXDRT.Sections
    Dim crxSection As CRAXDRT.Section
    Dim fi As Integer, fj As Integer
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
    cServer = fcnUTILITIES_GetPrivateProfileString("PLATFORM", "SERVER", App.Path & "\prjReportExporter.ini")
    cDataBase = fcnUTILITIES_GetPrivateProfileString("PLATFORM", "DATABASE", App.Path & "\prjReportExporter.ini")
    cUser = fcnUTILITIES_GetPrivateProfileString("PLATFORM", "USER", App.Path & "\prjReportExporter.ini")
    cPWD = fcnUTILITIES_GetPrivateProfileString("PLATFORM", "PWD", App.Path & "\prjReportExporter.ini")
    ' Open the report
    Set Report = application.OpenReport(str_in_ReportFileName, 1)
    Report.Database.Tables(1).SetLogOnInfo cServer, cDataBase, cUser, cPWD
    'This removes the schema from the Database Table's Location property.
    Set crxTables = Report.Database.Tables
    For Each crxTable In crxTables
        With crxTable
             .Location = .Name
        End With
    Next
    'Loop through the Report's Sections to find any subreports, and change them as well
    Set crxSections = Report.Sections
    For fi = 1 To crxSections.Count
        Set crxSection = crxSections(fi)
        For fj = 1 To crxSection.ReportObjects.Count
            If crxSection.ReportObjects(fj).Kind = crSubreportObject Then
                Set crxSubreportObject = crxSection.ReportObjects(fj)
                'Open the subreport, and treat like any other report
                Set crxSubReport = crxSubreportObject.OpenSubreport
                Set crxTables = crxSubReport.Database.Tables
                For Each crxTable In crxTables
                    With crxTable
                        .SetLogOnInfo cServer, cDataBase, cUser, cPWD
                        .Location = .Name
                    End With
                Next
            End If
        Next fj
    Next fi
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
    Set Report = Nothing
    Set application = Nothing
    Set crxTables = Nothing
    Set crxTable = Nothing
    Set crxSubreportObject = Nothing
    Set crxSubReport = Nothing
    Set crxSections = Nothing
    Set crxSection = Nothing
    Exit Sub
err_ExportReport:
    Err.Raise -12345123, "clsReportExporter.ExportReport", "Error: " & Err.Description
End Sub

Private Sub Class_Initialize()
    intCurrentParameter = 0
End Sub
