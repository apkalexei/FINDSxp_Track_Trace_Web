<%
	CheckError ("Error at the beginning of the page.")
	
	' ************************* SQL STATEMENT GENERATION FUNCTIONS *************************
	Function fcnSHIPGLANCE_ConstructSQLStatement(ByVal nNumParams, ByVal nTransType, ByVal sDateFrom, ByVal sDateTo)
	  Dim sSQL, sVal, lIndex, nPos
	  Dim bOpenGroup
	  Dim x

	  'Init SQL string
	  sSQL = ""

	  'Init Grouping flag
	  bOpenGroup = False
	  
	  'Construct SQL statement for parameters
	  For x = 0 To nNumParams - 1
	    'Check if Paramter value is not blank
	    If Len(Trim(txtField(x))) > 0 Then
	      'Check if this is not 1st parameter
	      If Len(sSQL) > 0 Then
	        'Concatenate Operator
	        sSQL = sSQL & " " & clbOperatorField(x - 1) & " "
	      End If
	      'Group OR'd parameters together to be evaluated 1st and add Parameters
	      sSQL = fcnSHIPGLANCE_GroupSQLStatements(nNumParams, sSQL, x, bOpenGroup)
	    End If
	  Next

	  'Enclose parameter string if necessary
	  If Len(sSQL) > 0 Then sSQL = "(" & sSQL & ")"

	  'Check if Date Range selected
	  If strDateRangeOn = "on" Then
	    'If there are parameters "AND" the Date Range
	    If Len(sSQL) > 0 Then
	      sSQL = sSQL & " AND "
	    End If
	    sSQL = sSQL & "(" & strDateType
	    sSQL = sSQL & " >= " & SQT & sDateFrom & " 12:00 AM" & SQT & ")"
	    sSQL = sSQL & " AND " & "(" & strDateType
	    sSQL = sSQL & " <= " & SQT & sDateTo & " 11:59 PM" & SQT & ")"
	  Else
	    'No parameters
	    If Len(sSQL) = 0 Then
	      sSQL = sSQL & "1=1"
	    End If
	  End If

	  sSQL = sSQL & " AND _hRequestUserIndex='" & session("UserIndex") & "'"
	
	  fcnSHIPGLANCE_ConstructSQLStatement = sSQL
	  CheckError ("Error at end of function fcnSHIPGLANCE_ConstructSQLStatement")
	End Function

	Function fcnSHIPGLANCE_GroupSQLStatements(ByVal nNumParams, ByVal sSQL, ByVal nCurrentIndex, bOpenGroup)
	  Dim sGrpMarker, x, sTmpSQL

	  'Init string
	  sGrpMarker = ""

	  'Loop through SQL string
	  For x = nCurrentIndex + 1 To nNumParams - 1
	    'Check if next Paramter value is not blank
	    If Len(Trim(txtField(x))) > 0 Then
	      'Check if next operator is "OR"
	      If clbOperatorField(x - 1) = "OR" Then
	        'If not already in Search
	        If Not bOpenGroup Then
	          'Start a grouping
	          sGrpMarker = "Open"
	          bOpenGroup = True
	        Else
	          'Continue the grouping
	          sGrpMarker = "Continue"
	        End If
	        Exit For
	      Else
	        'If in Search
	        If bOpenGroup Then
	          'End a grouping
	          sGrpMarker = "Close"
	          'End Search
	          bOpenGroup = False
	        End If
	        Exit For
	      End If
	    End If
	  Next

	  'If Closing a grouping or within an open grouping and no other non blanl parameters
	  If (sGrpMarker = "" And bOpenGroup) Or sGrpMarker = "Close" Then
	    'Construct SQL statement with parameter and ending Parand
	    sTmpSQL = sSQL & "(" & clbLUField(nCurrentIndex) & fcnSHIPGLANCE_ConstructSQLParam(nCurrentIndex) & "))"
	  'If Opening a grouping
	  ElseIf sGrpMarker = "Open" Then
	    'Construct SQL statement with parameter and starting Parand
	    sTmpSQL = sSQL & "((" & clbLUField(nCurrentIndex) & fcnSHIPGLANCE_ConstructSQLParam(nCurrentIndex) & ")"
	  'No Grouping, or continuing a group
	  ElseIf (sGrpMarker = "" And Not bOpenGroup) Or sGrpMarker = "Continue" Then
	    'Construct SQL statement with parameter
	    sTmpSQL = sSQL & "(" & clbLUField(nCurrentIndex) & fcnSHIPGLANCE_ConstructSQLParam(nCurrentIndex) & ")"
	  End If

	  fcnSHIPGLANCE_GroupSQLStatements = sTmpSQL
	  CheckError ("Error at end of function fcnSHIPGLANCE_GroupSQLStatements")
	End Function

	Function fcnSHIPGLANCE_ConstructSQLParam(ByVal nIndex)
	  Dim lIndex, sVal
	  Dim sParam

	  'Comment Field is anywhere in string or BLRider (MRK 1/21/08)
	  If (InStr(clbLUField(nIndex), "Comments") > 0) OR (InStr(clbLUField(nIndex), "BLRider") > 0) Then
	    sVal = Trim(txtField(nIndex))
	    sParam = WHERE_LIKE & SQT & WILDCARD & sVal & WILDCARD & SQT
	  Else
	    sVal = Trim(txtField(nIndex))
	    sParam = WHERE_LIKE & SQT & sVal & WILDCARD & SQT
	  End If

	  fcnSHIPGLANCE_ConstructSQLParam = sParam
	  CheckError ("Error at end of function fcnSHIPGLANCE_ConstructSQLParam")
	End Function
%>
	