VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWebReviewScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
' To create page:
'   1) Instantiate
'   2) Call getReviewScreenTable which returns the entire set of tables as a string value (in displayable html)

Private rs As Recordset
Private objConn As Connection

Private strOceanIndex As String
Private strUserRole As String
Private strUserIndex As String

Private strMenuTable As String

Private strTableStringData

Const HTML_TABLE_START = "<tr><td align=center><table bgcolor=#d0d0d0 border=1>"
Const HTML_TABLE_END = "</tr></table></td></tr>"
Const HTML_TITLE_START = "<tr><td height=15></td></tr><tr><td valign=top bgcolor=#b0b0b0><table width=100% border=0><tr><td align=left>&nbsp;<font size=4>"
Const HTML_TITLE_END = "</font></td><td align=right><a href='#screentop'><font size=1 color=blue>RETURN TO TOP&nbsp;</font></a></td></tr></table></td></tr><tr><td height=15></td></tr>"
Const HTML_1ROW_VALUE_START = "<tr><td valign=top><font size=2><b>"
Const HTML_1ROW_VALUE_MIDDLE = "&nbsp;</b></font></td><td><font face='arial, helvetica' size=1>"
Const HTML_1ROW_VALUE_END = "&nbsp;</font></td>"
Const HTML_TD_HEADING_START = "<td valign=top NOWRAP><font size=2><b>"
Const HTML_TD_HEADING_END = "</b></font></td>"
Const HTML_TD_VALUE_START = "<td valign=top><font face='arial, helvetica' size=1>"
Const HTML_TD_VALUE_END = "</font></td>"

Private Function getTDValue(str_in_Value) As String
    getTDValue = HTML_TD_VALUE_START & str_in_Value & "&nbsp;" & HTML_TD_VALUE_END
End Function

Private Function getTDHeading(str_in_Heading) As String
    getTDHeading = HTML_TD_HEADING_START & str_in_Heading & HTML_TD_HEADING_END
End Function

Private Function getTitle(str_in_Title As String) As String
    
    ' Create the title and the anchor tag
    getTitle = HTML_TITLE_START & "<a name=""" & Replace(str_in_Title, " ", "") & """>" & str_in_Title & HTML_TITLE_END

End Function

Private Function get1AddressRow(str_in_RSHeading As String, str_in_DBAddressType, str_in_DisplayAddressType As String) As String
    
  ' This function returns a row for an html table with the heading and the name, with a javascript hyperlink to open the address info window
  If Len("" & rs(str_in_RSHeading)) <> 0 And rs(str_in_RSHeading) <> "$$$HIDDEN$$$" Then
    get1AddressRow = HTML_1ROW_VALUE_START & str_in_DisplayAddressType & " Name " & HTML_1ROW_VALUE_MIDDLE & "<a href=""javascript:openWindow('" & strOceanIndex & "','" & rs(str_in_DBAddressType) & "','" & str_in_DisplayAddressType & "')""><font color=blue>" & rs(str_in_RSHeading) & "</font></a>" & HTML_1ROW_VALUE_END
  End If

End Function

Private Function get1ValueRow(str_in_Heading As String, str_in_RSHeading As String) As String
    
  'This function returns a row for an html table with the heading and the name, with a javascript hyperlink to open the address info window
  If Len("" & rs(str_in_RSHeading)) <> 0 And rs(str_in_RSHeading) <> "$$$HIDDEN$$$" Then
    get1ValueRow = HTML_1ROW_VALUE_START & str_in_Heading & " " & HTML_1ROW_VALUE_MIDDLE & rs(str_in_RSHeading) & HTML_1ROW_VALUE_END
  End If

End Function

Private Function getHTMLCustomerInfoTable() As String
    Dim strOutput As String
    strOutput = ""
    
    strOutput = strOutput & get1AddressRow("AcctName", "AcctType", "Account")
    strOutput = strOutput & get1AddressRow("ExporterName", "ExporterType", "Exporter")
    strOutput = strOutput & get1AddressRow("FwdrName", "ForwarderType", "Forwarder")
    strOutput = strOutput & get1ValueRow("Shippers Ref #:", "ShipperRefNo")
    strOutput = strOutput & get1ValueRow("Shippers PO #:", "ShipperPO")
    strOutput = strOutput & get1ValueRow("Forwarder Ref #:", "FwdrRefNo")
    strOutput = strOutput & get1ValueRow("Export Ref #:", "ExportRef")
    
    ' If NONE of the values were shown for this table, then don't even show the Table
    If strOutput <> "" Then
        strOutput = getTitle("Customer") & HTML_TABLE_START & strOutput & HTML_TABLE_END
        strMenuTable = strMenuTable & "<tr><td>&nbsp;<a href='#Customer'><FONT face='arial, helvetica' size=1 color=blue>CUSTOMER</font></a>&nbsp;</td></tr>"
    End If
    getHTMLCustomerInfoTable = strOutput
End Function

Private Function getHTMLReceiverInfoTable() As String
    Dim strOutput As String
    Dim strTemp As String
    
    strOutput = ""
    
    strOutput = strOutput & get1AddressRow("ConsigName", "ConsigType", "Consignee")
    strOutput = strOutput & get1AddressRow("NotifyName", "NotifyType", "Notify")
    strOutput = strOutput & get1AddressRow("Notify2Name", "Notify2Type", "Also Notify")
    strOutput = strOutput & get1AddressRow("UltConsigName", "UltConsigType", "Ultimate Consignee")
    
    If IsNull(rs("fPartiesRelated")) Or rs("fPartiesRelated") <> "$$$HIDDEN$$$" Then
        ' Display parties related as Yes/No
        strTemp = rs("fPartiesRelated")
        If strTemp = "True" Then
            strTemp = "YES"
        Else
            strTemp = "NO"
        End If
        
        strOutput = strOutput & HTML_1ROW_VALUE_START & "Parties Related " & HTML_1ROW_VALUE_MIDDLE & strTemp & HTML_1ROW_VALUE_END
    End If
    strOutput = strOutput & get1AddressRow("AgentName", "AgentType", "Recieving Agent")
    
    ' If NONE of the values were shown for this table, then don't even show the Table
    If strOutput <> "" Then
        strOutput = getTitle("Receiver") & HTML_TABLE_START & strOutput & HTML_TABLE_END
        strMenuTable = strMenuTable & "<tr><td>&nbsp;<a href='#Receiver'><FONT face='arial, helvetica' size=1 color=blue>RECEIVER</font></a>&nbsp;</td></tr>"
    End If
    getHTMLReceiverInfoTable = strOutput
End Function

Private Function getHTMLRouteInfoTable() As String
  Dim strOutput As String
    
  strOutput = ""
    
  strOutput = strOutput & get1ValueRow("Place of Receipt Description:", "PLRDescription")
  strOutput = strOutput & get1ValueRow("Port of Load Description:", "POLDescription")
  strOutput = strOutput & get1ValueRow("Port of Discharge Description:", "PODDescription")
  strOutput = strOutput & get1ValueRow("Transshipment Description:", "TShipDescription")
    
  strOutput = strOutput & get1ValueRow("Place of Delivery Description:", "PLDDescription")
  strOutput = strOutput & get1ValueRow("Country of Origin Description:", "CtryOriginDescription")
  strOutput = strOutput & get1ValueRow("State of Origin ID:", "State")
  strOutput = strOutput & get1ValueRow("Ultimate Destination Description:", "CtryUltDestDescription")

  ' If NONE of the values were shown for this table, then don't even show the Table
  If strOutput <> "" Then
    strOutput = getTitle("Route") & HTML_TABLE_START & strOutput & HTML_TABLE_END
    strMenuTable = strMenuTable & "<tr><td>&nbsp;<a href='#Route'><FONT face='arial, helvetica' size=1 color=blue>ROUTE</font></a>&nbsp;</td></tr>"
  End If
    
    getHTMLRouteInfoTable = strOutput
End Function

Private Function getHTMLOceanCarrierInfoTable() As String
    Dim strOutput As String
    strOutput = ""

    strOutput = strOutput & get1AddressRow("CarrierName", "CarrierType", "Carrier")
    strOutput = strOutput & get1ValueRow("Vessel Name:", "VesselName")
    strOutput = strOutput & get1ValueRow("Voyage:", "Voyage")
    strOutput = strOutput & get1ValueRow("Carrier Booking Number:", "CarrierBookingNo")
    strOutput = strOutput & get1ValueRow("Booking Mode:", "ServiceTypeID")
    strOutput = strOutput & get1ValueRow("Service Contranct ID:", "SvcContractID")
    strOutput = strOutput & get1ValueRow("Cutoff Date:", "CutoffDate")
    strOutput = strOutput & get1ValueRow("Sail Date:", "SailDate")
    strOutput = strOutput & get1ValueRow("OnBoard Date:", "OnBoardDate")
    strOutput = strOutput & get1ValueRow("TShip Date:", "TShipDate")
    strOutput = strOutput & get1ValueRow("ETA Date:", "ETADate")
    strOutput = strOutput & get1ValueRow("Actual Arrival Date:", "ActualArrivalDate")
    
    ' If NONE of the values were shown for this table, then don't even show the Table
    If strOutput <> "" Then
        strOutput = getTitle("Ocean Carrier") & HTML_TABLE_START & strOutput & HTML_TABLE_END
        strMenuTable = strMenuTable & "<tr><td>&nbsp;<a href='#OceanCarrier'><FONT face='arial, helvetica' size=1 color=blue>OCEAN CARRIER</font></a>&nbsp;</td></tr>"
    End If
    
    getHTMLOceanCarrierInfoTable = strOutput
End Function

Private Function getHTMLContainerPositionInfoTable() As String
    Dim strOutput As String
    strOutput = ""
    
    strOutput = strOutput & get1AddressRow("PickupTermName", "PickupTermType", "Pickup Terminal")
    strOutput = strOutput & get1AddressRow("DelivTermName", "DelivTermType", "Delivery Terminal")
    
    ' If NONE of the values were shown for this table, then don't even show the Table
    If strOutput <> "" Then
        strOutput = getTitle("Container Position") & HTML_TABLE_START & strOutput & HTML_TABLE_END
        strMenuTable = strMenuTable & "<tr><td>&nbsp;<a href='#ContainerPosition'><FONT face='arial, helvetica' size=1 color=blue>CONTAINER POSITION</font></a>&nbsp;</td></tr>"
    End If
    
    getHTMLContainerPositionInfoTable = strOutput
End Function

Private Function getHTMLFileInfoTable() As String
  Dim strOutput As String
    
  strOutput = ""
    
  strOutput = strOutput & get1ValueRow("Controller Name:", "ControllerName")
  strOutput = strOutput & get1ValueRow("Date Opened:", "DateOpened")
  strOutput = strOutput & get1ValueRow("Last Date Modified:", "LastDateModified")
    
  ' If NONE of the values were shown for this table, then don't even show the Table
  If strOutput <> "" Then
    strOutput = getTitle("File") & HTML_TABLE_START & strOutput & HTML_TABLE_END
    strMenuTable = strMenuTable & "<tr><td>&nbsp;<a href='#File'><FONT face='arial, helvetica' size=1 color=blue>FILE</font></a>&nbsp;</td></tr>"
  End If
    
  getHTMLFileInfoTable = strOutput

End Function
Private Function getHTMLTermInfoTable() As String
  Dim strOutput As String
    
  strOutput = ""
    
  strOutput = strOutput & get1ValueRow("Incoterm:", "IncoTermID")
  strOutput = strOutput & get1ValueRow("Payment Method:", "PymtMethod")
    
  ' If NONE of the values were shown for this table, then don't even show the Table
  If strOutput <> "" Then
    strOutput = getTitle("Terms") & HTML_TABLE_START & strOutput & HTML_TABLE_END
    strMenuTable = strMenuTable & "<tr><td>&nbsp;<a href='#Terms'><FONT face='arial, helvetica' size=1 color=blue>TERMS</font></a>&nbsp;</td></tr>"
  End If
    
  getHTMLTermInfoTable = strOutput

End Function

Private Function getHTMLLetterOfCreditInfoTable() As String
  Dim strOutput As String

  strOutput = ""

  strOutput = strOutput & get1ValueRow("L/C #:", "LCNum")
  strOutput = strOutput & get1ValueRow("Opening Bank Name:", "IssueBankName")
  strOutput = strOutput & get1ValueRow("Negotiating Bank Name:", "NegotiateBankName")
  strOutput = strOutput & get1ValueRow("Opening Date:", "IssueDate")
  strOutput = strOutput & get1ValueRow("Latest Ship Date:", "LatestSailDate")
  strOutput = strOutput & get1ValueRow("Expiration Date:", "ExpiryDate")
    
  ' If NONE of the values were shown for this table, then don't even show the Table
  If strOutput <> "" Then
    strOutput = getTitle("Letter Of Credit") & HTML_TABLE_START & strOutput & HTML_TABLE_END
    strMenuTable = strMenuTable & "<tr><td>&nbsp;<a href='#LetterOfCredit'><FONT face='arial, helvetica' size=1 color=blue>LETTER OF CREDIT</font></a>&nbsp;</td></tr>"
  End If

  getHTMLLetterOfCreditInfoTable = strOutput

End Function

Private Function getHTMLCargoInfoTable() As String
    Dim strOutput As String
    Dim rs2 As New ADODB.Recordset
    Dim strLastValue As String
    
    strOutput = ""
    
    ' Do not display this information for users that can't see it
    If rs("fDisplayCargoInfo") Then
        rs2.Open "fwpOceanCargoReview '" & strOceanIndex & "', '" & strUserIndex & "'", objConn
        
        strOutput = "<tr>"
        strOutput = strOutput & getTDHeading("Container #") & getTDHeading("Container Type") & getTDHeading("Commodity") & getTDHeading("LBS") & getTDHeading("KGS") & getTDHeading("CFT") & getTDHeading("CBM") & getTDHeading("Piece Count") & getTDHeading("Piece Type")
        strOutput = strOutput & "</tr>"

        Do While Not rs2.EOF
            strOutput = strOutput & "<tr>"
            'Only show container info for the first row which uses it
            If strLastValue <> rs2("OceanCargoIndex") Then
                Dim strTempTabNo As String
                ' For break bulk, there is no tab 0 to link to (there's no container info), so we must link to the 2nd tab
                If rs2("fBreakBulk") Then
                    strTempTabNo = "1"
                Else
                    strTempTabNo = "0"
                End If
                strOutput = strOutput & HTML_TD_VALUE_START & "<a href='cargodetail.asp?tab=" & strTempTabNo & "&cargoindex=" & rs2("OceanCargoIndex") & "&bb=" & rs2("fBreakBulk") & "'><font color=blue>" & rs2("ContainerNum") & "</font></a>" & "&nbsp;" & HTML_TD_VALUE_END & getTDValue(rs2("ContainerTypeID"))
            Else
                strOutput = strOutput & "<td>&nbsp;</td><td>&nbsp;</td>"
            End If
            strOutput = strOutput & getTDValue(rs2("UserCmdyID")) & getTDValue(rs2("GrossLBS")) & getTDValue(rs2("GrossKGS")) & getTDValue(rs2("CFT")) & getTDValue(rs2("CBM")) & getTDValue(rs2("PieceCount")) & getTDValue(rs2("PieceTypeID"))
            strLastValue = rs2("OceanCargoIndex")
            strOutput = strOutput & "</tr>"
            rs2.MoveNext
        Loop

        rs2.Close
        Set rs2 = Nothing
        
        ' If NONE of the values were shown for this table, then don't even show the Table
        If strOutput <> "" Then
            strOutput = getTitle("Cargo") & HTML_TABLE_START & strOutput & HTML_TABLE_END
            strMenuTable = strMenuTable & "<tr><td>&nbsp;<a href='#Cargo'><FONT face='arial, helvetica' size=1 color=blue>CARGO</font></a>&nbsp;</td></tr>"
        End If
    End If
    getHTMLCargoInfoTable = strOutput
End Function

Private Function getHTMLExportDeclarationInfoTable() As String
    Dim strOutput As String
    Dim rs2 As New ADODB.Recordset
    
    strOutput = ""
    'strOutput = strOutput & HTML_TABLE_START & "<tr><td>BK</td></tr>" & HTML_TABLE_END
    
    If rs("fDisplayExDecInfo") Then
        ' Create the table for AES STATUS
        strOutput = strOutput & get1ValueRow("AES Status:", "AESStatus") & HTML_TABLE_END & "<tr><td height=5></td></tr>" & HTML_TABLE_START
    
        rs2.Open "fwpExDecReview '" & strOceanIndex & "'", objConn
        
        strOutput = strOutput & "<tr>"
        strOutput = strOutput & getTDHeading("Sched. B #") & getTDHeading("QTY #1") & getTDHeading("Units") & getTDHeading("QTY #2") & getTDHeading("Units") & getTDHeading("Value") & getTDHeading("KGS") & getTDHeading("Info Code") & getTDHeading("Lic. Code") & getTDHeading("Class Code #")
        strOutput = strOutput & "</tr>"
        
        Do While Not rs2.EOF
            strOutput = strOutput & "<tr>"
            strOutput = strOutput & getTDValue(rs2("SchedBNum")) & getTDValue(rs2("PrimaryQtyValue")) & getTDValue(rs2("PrimaryQtyUnits")) & getTDValue(rs2("SecondaryQtyValue")) & getTDValue(rs2("SecondaryQtyUnits")) & getTDValue(rs2("GrossKGS")) & getTDValue(rs2("ExportInfoCode")) & getTDValue(rs2("ExportLicCode")) & getTDValue(rs2("ExportLicNo")) & getTDValue(rs2("ExportClassCodeNo"))
            strOutput = strOutput & "</tr>"
            rs2.MoveNext
        Loop
        
        rs2.Close
        Set rs2 = Nothing
        
        ' If NONE of the values were shown for this table, then don't even show the Table
        If strOutput <> "" Then
            strOutput = getTitle("Export Declaration") & HTML_TABLE_START & strOutput & HTML_TABLE_END
            strMenuTable = strMenuTable & "<tr><td>&nbsp;<a href='#ExportDeclaration'><FONT face='arial, helvetica' size=1 color=blue>EXPORT DECLARATION</font></a>&nbsp;</td></tr>"
        End If
    End If
    
    getHTMLExportDeclarationInfoTable = strOutput
End Function

Public Function getJavaScriptTableDataCode() As String
    getJavaScriptTableDataCode = strMenuTable
End Function

Public Function getReviewScreenTable() As String
    getReviewScreenTable = strTableStringData
End Function

Public Sub prepareReviewScreenData(ByVal str_in_DBConnectionString As String, ByVal int_in_FGID As Long, ByVal str_in_UserRole As String, ByVal str_in_UserIndex As String)
    On Error GoTo ERR_prepareReviewScreenData

    ' Keep track of where I'm at in the code for debuggin' -- error handler passes the error to ASP page
    Dim strLocation As String
    Dim strTemp As String
    Dim strSPCall As String
    Set objConn = New ADODB.Connection
    Set rs = New ADODB.Recordset

    strLocation = "Start"
    
    strLocation = "DB Connect"
    objConn.Open str_in_DBConnectionString
    
    ' Open the recordset
    strSPCall = "fwpOceanHeaderDetail " & int_in_FGID & ", '" & str_in_UserIndex & "'"
    strLocation = "RS Fetch (calling """ & strSPCall & """) "
    rs.Open strSPCall, objConn
    
    strOceanIndex = int_in_FGID
    strUserRole = str_in_UserRole
    strUserIndex = str_in_UserIndex
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        strTemp = "<table border=0 cellpadding=0 cellspacing=0>"
        strLocation = "Before table generation"
        strLocation = "Customer Table Generation"
        strTemp = strTemp & getHTMLCustomerInfoTable()
        strLocation = "Receiver Table Generation"
        strTemp = strTemp & getHTMLReceiverInfoTable()
        strLocation = "Route Table Generation"
        strTemp = strTemp & getHTMLRouteInfoTable()
        strLocation = "Ocean Carrier Table Generation"
        strTemp = strTemp & getHTMLOceanCarrierInfoTable()
        strLocation = "Container Position Table Generation"
        strTemp = strTemp & getHTMLContainerPositionInfoTable()
        strLocation = "Term Table Generation"
        strTemp = strTemp & getHTMLTermInfoTable()
        strLocation = "File Table Generation"
        strTemp = strTemp & getHTMLFileInfoTable()
        strLocation = "LOC Table Generation"
        strTemp = strTemp & getHTMLLetterOfCreditInfoTable()
        strLocation = "Cargo Table Generation"
        strTemp = strTemp & getHTMLCargoInfoTable()
        strLocation = "ExDec Table Generation"
        strTemp = strTemp & getHTMLExportDeclarationInfoTable()
        
    
        strTemp = strTemp & "</table>"
    Else
        strLocation = "Calling """ & strSPCall & """ returned no records."
        GoTo ERR_prepareReviewScreenData
    End If
    
    'Return the html page
    strTableStringData = strTemp
    Exit Sub
ERR_prepareReviewScreenData:
    Call Err.Raise(-999123, "prepareReviewScreenData", "Error in function 'prepareReviewScreenData' of component. " & Err.Number & ", " & Err.Description & ".  Location: " & strLocation)
End Sub
