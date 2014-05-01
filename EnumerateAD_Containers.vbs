'********************************************************************
'* Title: Active Directory Tools - Export all user accounts for an OU
'* Author: Peter Warren
'* Last modified by: Peter Warren
'* Last modified date: 1-May-2014
'*
'********************************************************************

Option Explicit

'********************************************************************
'* Declare Global Constants - Scripts default parameters
'********************************************************************
Const defaultAgeAccountSuspend = "60"
Const defaultAgeAccountRemoval = "90"


'Const defaultScope = "1"	'Target OU only
Const defaultScope = "2"	'Target OU and all child OUs

'Const defaultOU = "OU=Users,DC=xxx,DC=com"



'********************************************************************
'* Declare Global Variables - WMI Object Management
'********************************************************************
Dim strDN
Dim strScope
Dim strType
Dim oRootDSE


'********************************************************************
'* Declare Global Variables - XML Document Object Model (DOM)
'********************************************************************
Dim xDoc
	Dim xIntro ' Document metadata e.g. 'xml version"1.0"'
	Dim xRoot
		Dim xSummary
			Dim xSummaryTitle
			Dim xSummaryParameters
				Dim xParaAgeAccountSuspend
				Dim xParaAgeAccountRemoval
				Dim xParaType
				Dim xParaScope
				Dim xParaOU
			Dim xSummaryStarttime
			Dim xSummaryEndtime
			Dim xSummaryElapsetime
		Dim xRecords
			Dim xRecord
		Dim xErrors
			Dim xError
				Dim xErrorPara
				Dim xErrorNum
				Dim xErrorSrc
				Dim xErrorDesc

'********************************************************************
'* Declare Global Variables - XML DOM Elements
'********************************************************************
'************* Attribute of: Computer, Group and User *************
Dim xObjectClass
Dim xdistinguishedName




'************* Attribute of: Container *************
Dim xC
Dim xCo
Dim xCountryCode
Dim xDescription
Dim xName

Dim xObjectCategory
Dim xObjectGUID
Dim xOu

Dim xUrl
Dim xWhenCreated
Dim xWhenChanged
Dim xGPLink
Dim xGPOptions






'********************************************************************
'* Function: ExtractCommon_OpenLDAP(strDN, strFilter)
'* Purpose: 
'* Input:   strDN		:	Distinguished Name as a string
'*			strFilter	:	Object type to be filtered on as a string (e.g. Computer, Group, or User)
'* Output:  Object
'* Notes:  
'*
Function ExtractCommon_OpenLDAP(strDN, strFilter)
	Set ExtractCommon_OpenLDAP = GetObject("LDAP://" & strDN)
	ExtractCommon_OpenLDAP.Filter = Array(strFilter)
End Function
'*
'********************************************************************


'********************************************************************
'* Function: CreateXMLDOMElement(sDescription, objAttribute, xParent)
'* Purpose: 
'* Input:	sDescription	:
'* 			objAttribute	:
'* 			xParent			:
'* Output:  Object
'* Notes:  
'*
Function CreateXMLDOMElement(sDescription, objParent, xParent)
	Dim objAttribValue

	On Error Resume Next
	objAttribValue = objParent.get(sDescription)

	If Err.Number = 0 Then
		Set CreateXMLDOMElement = xDoc.createElement(sDescription)
		CreateXMLDOMElement.text = objAttribValue
		xParent.appendChild CreateXMLDOMElement
	Else
		Set CreateXMLDOMElement = Nothing
		Err.Clear
	End if
	On Error Goto 0
End Function



'********************************************************************
'* Function: CreateXMLDOMElementFromString(sDescription, sAttribute, xParent)
Function CreateXMLDOMElementFromString(sDescription, sAttribute, xParent)
	Set CreateXMLDOMElementFromString = xDoc.createElement(sDescription)
	CreateXMLDOMElementFromString.text = sAttribute
	xParent.appendChild CreateXMLDOMElementFromString
End Function


'********************************************************************
'* Function: xxxBoolean(?, ?)
' Try using Case to improve speed
'Function xxxBoolean(sDescription, objAttribute, ObjComparator, xCompareSmtp2Upn)'
'If IsEmpty(objAttribute) = TRUE The'n
'	Set xxxBoolean = Nothing
'Else
'	Set xxxBoolean = xDoc.createElement(sDescription)
'	If objAttribute AND ObjComparator Then
'		xxxBoolean.Text = "TRUE"
'	Else 
'		xxxBoolean.Text = "FALSE"
'	End If			
'	xCompareSmtp2Upn.appendChild xxxBoolean
'End if
''d Function


'********************************************************************
'* Function:	ExtractCommon_IntegerToDate(ByVal intDateEpoch)
'* Purpose:		Function to convert Integer value to a date, not adjustment is made for local time zone bias
'* Input:		intDateEpoch	:	Integer which represents a date as seconds from epoch
'* Output:		Date
'* Notes:  
'*
Function ExtractCommon_IntegerToDate(ByVal intDateEpoch)
	ExtractCommon_IntegerToDate = CDate(intDateEpoch) + #1/1/1601#
End Function
'*
'********************************************************************


'********************************************************************
'* Function:	ExtractCommon_Integer8ToInteger(ByVal objInteger8)
'* Purpose:	Function to convert Integer8 (64-bit) value from an object to an Integer
'* Input:   intDateEpoch	:	Integer which represents a date as seconds from epoch
'* Output:  Date
'* Notes:  
'*
Function ExtractCommon_Integer8ToInteger(ByVal objInteger8)
	Dim intInteger8
	Dim intHighPart:intHighPart = objInteger8.HighPart
	Dim intLowPart:intLowPart = objInteger8.LowPart
	
	If (intLowPart < 0) Then
		intHighPart = intHighPart + 1
	End If
	
	intInteger8 = intHighPart * (2^32) + intLowPart 
	intInteger8 = intInteger8 / (60 * 10000000)
	intInteger8 = intInteger8 / 1440
	ExtractCommon_Integer8ToInteger = intInteger8
End Function
'*
'********************************************************************


'********************************************************************
'* Sub ExtractObject_ExportRecursive
'* Purpose:
'* Input:   strZoneOU
'* Output:  None
'* Notes:  
'*
Sub ExtractObject_ExportRecursive (ByVal strZoneOU)
	Dim objZoneOU
	Dim objZoneChildOU
	
	Set objZoneOU = ExtractCommon_OpenLDAP(strZoneOU, "organizationalUnit")
	ExtractObject_ExportSite(objZoneOU.distinguishedName)
	
	For Each objZoneChildOU In objZoneOU
		ExtractObject_ExportRecursive (objZoneChildOU.distinguishedName)
	Next
End Sub
'*
'********************************************************************


'********************************************************************
'* Sub ExtractObject_ExportSite
'* Purpose: Connects to site OU via LDAP and extracts all objects of specified type.
'* The Sub then populates Excel with a set of attributes for each object found
'* Input:   strSiteOU
'* Output:  
'* Notes:  
'*
'********************************************************************
Sub ExtractObject_ExportSite (ByVal strSiteOU)
	Dim objSiteOU
'	Dim objSiteChild
'	Dim objAccountType
'	Dim intLastLogonTimestamp
'	Dim intLockoutTime
'	Dim intWhenCreated
'	Dim intPwdLastSet
'	Dim memberCounter
'	Dim memberOfCounter
'	Dim objMember
'	Dim objMemberOf
'	Dim sPwdLastSet
	Dim intDateDiffCheck
	
	Dim colMembers
	Dim sGroupMember

	Dim colMembersOf
	Dim sGroupMemberOf
	
	
	Set objSiteOU = ExtractCommon_OpenLDAP(strSiteOU, strType)
'	For Each objSiteChild In objSiteOU
'		objAccountType = objSiteChild.class
'		If strType <> objAccountType Then
'			Exit For
'		End If

		'* Creates the XML Document object which will contain all elements for the current AD object
		Set xRecord = xDoc.createElement("Record")
		xRecords.appendChild xRecord


'        Set xOu = xDoc.createElement("OU")
'		xOu.text = objSiteOU.name
'		xOu.text = objSiteOU.get("name")
'		xRecord.appendChild xOu

		Dim xObjectSID
		Dim objNtSecurityDescriptor
		
        Set xObjectSID = xDoc.createElement("nTSecurityDescriptor")
		Set objNtSecurityDescriptor = objSiteOU.get("nTSecurityDescriptor")
		xObjectSID.text = objNtSecurityDescriptor.Control 
		xRecord.appendChild xObjectSID
		 
		
		'********************************************************************
		'* Extracts attributes common to all object types
        Set xC = CreateXMLDOMElement("C", objSiteOU, xRecord) ' Enumerate
        Set xCo = CreateXMLDOMElement("Co", objSiteOU, xRecord) ' Enumerate
        Set xDescription = CreateXMLDOMElement("Description", objSiteOU, xRecord) ' Enumerate
        Set xName = CreateXMLDOMElement("Name", objSiteOU, xRecord) ' Enumerate
        Set xObjectCategory = CreateXMLDOMElement("ObjectCategory", objSiteOU, xRecord) ' Enumerate
        'Set xObjectClass = CreateXMLDOMElement("ObjectClass", objSiteOU, xRecord) ' Enumerate
        'Set xObjectGUID = CreateXMLDOMElement("ObjectGUID", objSiteOU, xRecord) ' Enumerate

        Set xOu = CreateXMLDOMElement("Ou", objSiteOU, xRecord) ' Enumerate
        Set xUrl = CreateXMLDOMElement("Url", objSiteOU, xRecord) ' Enumerate
        Set xWhenCreated = CreateXMLDOMElement("WhenCreated", objSiteOU, xRecord) ' Enumerate
        Set xWhenCreated = CreateXMLDOMElement("WhenCreated", objSiteOU, xRecord) ' Enumerate

        Set xdistinguishedName = CreateXMLDOMElement("distinguishedName", objSiteOU, xRecord) ' Enumerate
        Set xGPLink = CreateXMLDOMElement("GPLink", objSiteOU, xRecord) ' Enumerate
        Set xGPOptions = CreateXMLDOMElement("xGPOption", objSiteOU, xRecord) ' Enumerate


		Set xObjectSID = CreateXMLDOMElement("nTSecurityDescriptor", objSiteOU, xRecord) ' Enumerate


		'********************************************************************

'		Next
End Sub


'********************************************************************
'* Sub Main()
'* Purpose: Main component of Active Directory Extraction script.
'* Input:   
'* Output:  
'*
'********************************************************************
Sub Main()

'********************************************************************
'* Setup variables to calculate time taken to execute script
Dim startTime:	startTime = Now
Dim stopTime
Dim elapsedTime
'********************************************************************

Dim objArgs
Set objArgs = WScript.Arguments

	'********************************************************************
	'* InputBox to select object type to be extracted from AD. Includes input validation
'Wscript.Echo "Argument count: " & WScript.Arguments.Count

    If WScript.Arguments.Count = 2 Then
        '	strType = objArgs(0)
        strScope = objArgs(0)
       strDN = objArgs(1)

        'Select Case strType
        '    Case "1"
        '        strType = ADS_OBJ_TYPE_COMPUTER
        '    Case "2"
        '        strType = ADS_OBJ_TYPE_GROUP
        '    Case "3"
        '        strType = ADS_OBJ_TYPE_USER
        'End Select

    Else
        'strType = InputBox("Enter the type object to extract from Active Directory Computer[1], Group[2], or User[3] ", "Input Object Type", defaultType)
        'strType = defaultType
        '	If strType = "1" OR strType = "2" OR strType = "3" Then




        'Select Case strType
        '    Case "1"
        '        strType = ADS_OBJ_TYPE_COMPUTER
        '    Case "2"
        '        strType = ADS_OBJ_TYPE_GROUP
        '    Case "3"
        '        strType = ADS_OBJ_TYPE_USER
        'End Select
        '********************************************************************

        '********************************************************************
        '* InputBox to select mode of operation. Extract data from: Single OU or Recursively 
        strScope = InputBox("Enter the mode of operation. Single OU Only  [1] or additionally query child OU [2]", "Input Scope", defaultScope)
        '	strScope = defaultScope
        '		If strScope = "1" OR strScope = "2" Then
        '* InputBox to select the DN of the target OU
        strDN = InputBox("Enter the distinguished name of a Site container", "Input Site OU", defaultOU)
        '		strDN = defaultOU
        '********************************************************************

    End If
'********************************************************************
'* Create XML Document Object Model
		Set xDoc = CreateObject("Microsoft.XMLDOM")
		
		Set xRoot = xDoc.createElement("Document_Root")
		xDoc.appendChild xRoot
		
		Set xSummary = xDoc.createElement("Summary")
		xRoot.appendChild xSummary
		
		Set xRecords = xDoc.createElement("Records")
		xRoot.appendChild xRecords

		Set xErrors = xDoc.createElement("Records")
		xRoot.appendChild xErrors

'********************************************************************

'********************************************************************
'* Execute either Active Directory extract in either single or recursive mode
		Select Case strScope
		Case "1"
			ExtractObject_ExportSite(strDN)
		Case "2"
			ExtractObject_ExportRecursive(strDN)
		End Select
'********************************************************************

'	End If
'End If


'********************************************************************
'* Setup variables to calculate time taken to execute script
stopTime = Now
elapsedTime = DateDiff("s",startTime,stopTime)
'********************************************************************


'********************************************************************
'* Update XML header and summary. Write XML DOM to File
Set xIntro = xDoc.createProcessingInstruction("xml","version='1.0'")
xDoc.insertBefore xIntro,xDoc.childNodes(0)

Set xSummaryTitle = xDoc.createElement("Title")
xSummaryTitle.Text = "Under construction!!!"
xSummary.appendChild xSummaryTitle

Set xSummaryParameters = xDoc.createElement("Parameters")
xSummary.appendChild xSummaryParameters

	Set xParaAgeAccountSuspend = xDoc.createElement("AgeAccountSuspend")
	xParaAgeAccountSuspend.Text = defaultAgeAccountSuspend
	xSummaryParameters.appendChild xParaAgeAccountSuspend

	Set xParaAgeAccountRemoval = xDoc.createElement("AgeAccountRemoval")
	xParaAgeAccountRemoval.Text = defaultAgeAccountRemoval
	xSummaryParameters.appendChild xParaAgeAccountRemoval

	Set xParaType = xDoc.createElement("Type")
	xParaType.Text = strType
	xSummaryParameters.appendChild xParaType

	Set xParaScope = xDoc.createElement("Scope")
	xParaScope.Text = strScope
	xSummaryParameters.appendChild xParaScope

	Set xParaOU = xDoc.createElement("OU")
	xParaOU.Text = strDN
	xSummaryParameters.appendChild xParaOU


Set xSummaryStarttime = xDoc.createElement("Start_Time")
xSummaryStarttime.Text = startTime
xSummary.appendChild xSummaryStarttime

Set xSummaryEndtime = xDoc.createElement("End_Time")
xSummaryEndtime.Text = stopTime
xSummary.appendChild xSummaryEndtime

Set xSummaryElapsetime = xDoc.createElement("Elapse_Time")
xSummaryElapsetime.Text = elapsedTime
xSummary.appendChild xSummaryElapsetime

Dim OutputFile 
OutputFile = "AD_" & strType & Day(Now) & MonthName(Month(Now),True) & Year(Now) & ".xml"
xDoc.Save OutputFile
'********************************************************************
Wscript.Echo "Script Completed in : " & elapsedTime & " seconds"


End Sub

Main

'********************************************************************
'*                                                                  *
'*                           End of File                            *
'*                                                                  *
'********************************************************************









