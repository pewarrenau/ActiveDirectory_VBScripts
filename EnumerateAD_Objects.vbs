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

'Const enumerateGroupMembers = TRUE
'Const enumerateGroupMembers = FALSE

'Const defaultType = "1"		'Computer
'Const defaultType = "2"		'Group
Const defaultType = "3"		'User

'Const defaultScope = "1"	'Target OU only
Const defaultScope = "2"	'Target OU and all child OUs

Const defaultOU = "OU=Users,DC=xxx,DC=com"



'********************************************************************
'* Declare Global Constants - Active Directory Management
'********************************************************************
Const ADS_OBJ_TYPE_COMPUTER = "computer"
Const ADS_OBJ_TYPE_GROUP = "group"
Const ADS_OBJ_TYPE_USER = "user"

'*************** User Account Control Flags ****************
Const ADS_UF_SCRIPT = &H1
Const ADS_UF_ACCOUNTDISABLE = &H2
Const ADS_UF_HOMEDIR_REQUIRED = &H8
Const ADS_UF_LOCKOUT = &H10
Const ADS_UF_PASSWD_NOTREQD = &H20
Const ADS_UF_PASSWD_CANT_CHANGE = &H40
Const ADS_UF_ENCRYPTED_TEXT_PWD_ALLOWED = &H80
Const ADS_UF_TEMP_DUPLICATE_ACCOUNT = &H100
Const ADS_UF_NORMAL_ACCOUNT = &H200
Const ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = &H800
Const ADS_UF_WORKSTATION_TRUST_ACCOUNT = &H1000
Const ADS_UF_SERVER_TRUST_ACCOUNT = &H2000
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
Const ADS_UF_MNS_LOGON_ACCOUNT = &H20000
Const ADS_UF_SMARTCARD_REQUIRED = &H40000
Const ADS_UF_TRUSTED_FOR_DELEGATION = &H80000
Const ADS_UF_NOT_DELEGATED = &H100000
Const ADS_UF_USE_DES_KEY_ONLY = &H200000
Const ADS_UF_DONT_REQ_PREAUTH = &H400000
Const ADS_UF_PASSWORD_EXPIRED = &H800000
Const ADS_UF_TRUSTED_TO_AUTH_FOR_DELEGATION = &H1000000
Const ADS_UF_PARTIAL_SECRETS_ACCOUNT = &H4000000

'******************* Group Type / Scope ********************
Const ADS_UF_SYSTEM = &H1
Const ADS_UF_SCOPE_GLOBAL = &H2
Const ADS_UF_SCOPE_LOCAL = &H4
Const ADS_UF_SCOPE_UNIVERSAL = &H8
Const ADS_UF_SECURITY_GROUP = &H80000000

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

Dim xOu
Dim xName
Dim xCn
Dim xWhenCreated
Dim xWhenChanged
Dim xManagedBy
Dim xDescription
Dim xCountry
Dim xCountryCode


'************* Attribute of: Computer and User *************
Dim xDisplayName
Dim xMsDSUserAccountControlComputed

Dim xPwdLastSet
Dim xPwdMuchChange
Dim xLastLogonTimestamp
Dim xLockoutTime

'***************** Attribute of: Computer ******************
Dim xOs
'Dim xOS:			Const descriptionsOS = "Operatingsystem"
Dim xOsVersion
Dim xOsServicePack
Dim xOsHotFix
'******************* Attribute of: Group *******************
Dim xMembers
	Dim xMember
	Dim xMembersElement
	Dim xMembersCount
Dim xMembersOf
	Dim xMembersOfElement
	Dim xMemberOfCount


Dim xGroupType
	Dim xGroupTypeValue
	Dim xGroupCreatedBySystem
	Dim xGroupTypeStr
	Dim xGroupScope
'************** Attribute of: Group and User ***************
Dim xEmail
'******************* Attribute of: User ********************
Dim xUserPrincipalName
Dim xGivenName
Dim xSn
Dim xDepartment
Dim xHomeDirectory
Dim xHomeDrive
Dim xProfilePath
Dim xCompareSmtp2Upn
Dim xTelephoneNumber
Dim xMobile
Dim xFacsimileTelephoneNumber
Dim xIpPhone
Dim xStreetAddress
Dim xL
Dim xSt
Dim xCo
Dim xPostalCode
Dim xManager
Dim xUserWorkstations

'******************* Attribute of: Exchange ********************
Dim xHomeMDB
Dim xHomeMTA
Dim xLegacyExchangeDN
Dim xMDBOverQuotaLimit
Dim xMDBStorageQuota
Dim xMDBUseDefaults
Dim xMsExchHideFromAddressLists
Dim xMsExchHomeServerName
Dim xMsExchMobileMailboxFlags
Dim xMsExchRecipientDisplayType
Dim xMsExchRecipientTypeDetails



'******************* Attribute of: Office Communicator 2007********************
Dim xMsRTCSIP
	Dim xMsRTCSIPFederationEnabled
	Dim xMsRTCSIPInternetAccessEnabled
	Dim xMsRTCSIPOptionFlags
	Dim xMsRTCSIPPrimaryHomeServer
	Dim xMsRTCSIPPrimaryUserAddress
	Dim xMsRTCSIPUserEnabled


'******************* Attribute of: UserAccountControl ********************
Dim xUserAccountControl
	Dim xUserAccountControlValue
	Dim xUACScript
	Dim xUACAccountDisabled
	Dim xUACHomedirRequired
	Dim xUACPasswdNotreqd
	Dim xUACPasswdCantChange
	Dim xUACEncryptTextPwdAllowed
	Dim xUACTempDuplicateAccount
	Dim xUACNormalAccount
	Dim xUACInterdomainTrustAccount
	Dim xUACWorkstationTrustAccount
	Dim xUACServerTrustAccount
	Dim xUACDontExpirePasswd
	Dim xUACMnsLogonAccount
	Dim xUACSmartcardRequired
	Dim xUACTrustedForDelegation
	Dim xUACNotDelegation
	Dim xUACUseDesKeyOnly
	Dim xUACDontReqPreauth
	Dim xUACPasswordExpired
	Dim xUACTrustedToAuthForDelegation
	Dim xUACPartialSecretsAccount


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
Function xxxBoolean(sDescription, objAttribute, ObjComparator, xCompareSmtp2Upn)
	If IsEmpty(objAttribute) = TRUE Then
		Set xxxBoolean = Nothing
	Else
		Set xxxBoolean = xDoc.createElement(sDescription)
		If objAttribute AND ObjComparator Then
			xxxBoolean.Text = "TRUE"
		Else 
			xxxBoolean.Text = "FALSE"
		End If			
		xCompareSmtp2Upn.appendChild xxxBoolean
	End if
End Function


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
	Dim objSiteChild
	Dim objAccountType
	Dim intLastLogonTimestamp
	Dim intLockoutTime
	Dim intWhenCreated
	Dim intPwdLastSet
	Dim memberCounter
	Dim memberOfCounter
	Dim objMember
	Dim objMemberOf
	Dim sPwdLastSet
	Dim intDateDiffCheck
	
	Dim colMembers
	Dim sGroupMember

	Dim colMembersOf
	Dim sGroupMemberOf
	
	
	Set objSiteOU = ExtractCommon_OpenLDAP(strSiteOU, strType)
	For Each objSiteChild In objSiteOU
		objAccountType = objSiteChild.class
		If strType <> objAccountType Then
			Exit For
		End If

		'* Creates the XML Document object which will contain all elements for the current AD object
		Set xRecord = xDoc.createElement("Record")
		xRecords.appendChild xRecord
	
		'********************************************************************
		'* Extracts attributes common to all object types
		Set xObjectClass = CreateXMLDOMElement("objectClass", objSiteChild, xRecord) ' Enumerate
		Set xDistinguishedName = CreateXMLDOMElement("distinguishedName", objSiteChild, xRecord)
		Set xName = CreateXMLDOMElement("name", objSiteChild, xRecord)
		Set xCn = CreateXMLDOMElement("cn", objSiteChild, xRecord)
		Set xCountry = CreateXMLDOMElement("co", objSiteChild, xRecord)
		Set xCountryCode = CreateXMLDOMElement("c", objSiteChild, xRecord)
		Set xWhenCreated = CreateXMLDOMElement("whenCreated", objSiteChild, xRecord)
		Set xWhenChanged = CreateXMLDOMElement("whenChanged", objSiteChild, xRecord)
		Set xDescription = CreateXMLDOMElement("description", objSiteChild, xRecord)
		Set xManagedBy = CreateXMLDOMElement("managedBy", objSiteChild, xRecord)
		Set xDisplayName = CreateXMLDOMElement("displayName", objSiteChild, xRecord)
		'********************************************************************

		'********************************************************************
		'* Extracts attributes common to computer all objects
		Set xOs = CreateXMLDOMElement("operatingSystem", objSiteChild, xRecord)
		Set xOsVersion = CreateXMLDOMElement("operatingSystemVersion", objSiteChild, xRecord)
		Set xOsServicePack = CreateXMLDOMElement("operatingSystemServicePack", objSiteChild, xRecord)
		Set xOsHotFix = CreateXMLDOMElement("operatingSystemHotfix", objSiteChild, xRecord)
		'********************************************************************
		
		'********************************************************************
		'* Extracts attributes common to user all objects
		Set xMsExchHideFromAddressLists = CreateXMLDOMElement("msExchHideFromAddressLists", objSiteChild, xRecord)
		Set xUserPrincipalName = CreateXMLDOMElement("UserPrincipalName", objSiteChild, xRecord)
		Set xGivenName = CreateXMLDOMElement("givenName", objSiteChild, xRecord)
		Set xSn = CreateXMLDOMElement("sn", objSiteChild, xRecord)
		Set xDepartment = CreateXMLDOMElement("department", objSiteChild, xRecord)
		Set xHomeDirectory = CreateXMLDOMElement("homeDirectory", objSiteChild, xRecord)
		Set xHomeDrive = CreateXMLDOMElement("homeDrive", objSiteChild, xRecord)
		Set xProfilePath = CreateXMLDOMElement("profilePath", objSiteChild, xRecord)
		Set xManager = CreateXMLDOMElement("manager", objSiteChild, xRecord)
		Set xTelephoneNumber = CreateXMLDOMElement("telephoneNumber", objSiteChild, xRecord)
		Set xMobile = CreateXMLDOMElement("mobile", objSiteChild, xRecord)
		Set xFacsimileTelephoneNumber = CreateXMLDOMElement("facsimileTelephoneNumber", objSiteChild, xRecord)
		Set xIpPhone = CreateXMLDOMElement("ipPhone", objSiteChild, xRecord)
		Set xStreetAddress = CreateXMLDOMElement("streetAddress", objSiteChild, xRecord)
		Set xL = CreateXMLDOMElement("l", objSiteChild, xRecord)
		Set xSt = CreateXMLDOMElement("st", objSiteChild, xRecord)
		Set xCo = CreateXMLDOMElement("co", objSiteChild, xRecord)
		Set xPostalCode = CreateXMLDOMElement("postalCode", objSiteChild, xRecord)

		Set xUserWorkstations = CreateXMLDOMElement("userWorkstations", objSiteChild, xRecord)
		
		'********************************************************************


'		Set xMsDSUserAccountControlComputed = CreateXMLDOMElement("ms_DSUserAccountControlComputed", objSiteChild.ms-DS-User-Account-Control-Computed, xRecord)	'New field used in 2003 Domains
		'********************************************************************
		'* This section reports the UAC and enumerates all UAC flags elements
		If isEmpty(objSiteChild.userAccountControl) = FALSE Then
			Set xUserAccountControl = xDoc.createElement("UserAccountControl")
			xRecord.appendChild xUserAccountControl
			'* Following code generates all flags derived from the UserAccountControl value
			Set xUserAccountControlValue = CreateXMLDOMElement("userAccountControl", objSiteChild, xUserAccountControl)
			Set xUACScript = xxxBoolean("ADS_UF_SCRIPT", objSiteChild.userAccountControl, ADS_UF_SCRIPT, xUserAccountControl)
			Set xUACAccountDisabled = xxxBoolean("ADS_UF_ACCOUNTDISABLE", objSiteChild.userAccountControl, ADS_UF_ACCOUNTDISABLE, xUserAccountControl)
			Set xUACHomedirRequired = xxxBoolean("ADS_UF_HOMEDIR_REQUIRED", objSiteChild.userAccountControl, ADS_UF_HOMEDIR_REQUIRED, xUserAccountControl)
			Set xUACPasswdNotreqd = xxxBoolean("ADS_UF_PASSWD_NOTREQD", objSiteChild.userAccountControl, ADS_UF_PASSWD_NOTREQD, xUserAccountControl)
			Set xUACPasswdCantChange = xxxBoolean("ADS_UF_PASSWD_CANT_CHANGE", objSiteChild.userAccountControl, ADS_UF_PASSWD_CANT_CHANGE, xUserAccountControl)
			Set xUACEncryptTextPwdAllowed = xxxBoolean("ADS_UF_ENCRYPTED_TEXT_PWD_ALLOWED", objSiteChild.userAccountControl, ADS_UF_ENCRYPTED_TEXT_PWD_ALLOWED, xUserAccountControl)
			Set xUACTempDuplicateAccount = xxxBoolean("ADS_UF_TEMP_DUPLICATE_ACCOUNT", objSiteChild.userAccountControl, ADS_UF_TEMP_DUPLICATE_ACCOUNT, xUserAccountControl)
			Set xUACNormalAccount = xxxBoolean("ADS_UF_NORMAL_ACCOUNT", objSiteChild.userAccountControl, ADS_UF_NORMAL_ACCOUNT, xUserAccountControl)
			Set xUACInterdomainTrustAccount = xxxBoolean("ADS_UF_INTERDOMAIN_TRUST_ACCOUNT", objSiteChild.userAccountControl, ADS_UF_INTERDOMAIN_TRUST_ACCOUNT, xUserAccountControl)
			Set xUACWorkstationTrustAccount = xxxBoolean("ADS_UF_WORKSTATION_TRUST_ACCOUNT", objSiteChild.userAccountControl, ADS_UF_WORKSTATION_TRUST_ACCOUNT, xUserAccountControl)
			Set xUACServerTrustAccount = xxxBoolean("ADS_UF_SERVER_TRUST_ACCOUNT", objSiteChild.userAccountControl, ADS_UF_SERVER_TRUST_ACCOUNT, xUserAccountControl)
			Set xUACDontExpirePasswd = xxxBoolean("ADS_UF_DONT_EXPIRE_PASSWD", objSiteChild.userAccountControl, ADS_UF_DONT_EXPIRE_PASSWD, xUserAccountControl)
			Set xUACMnsLogonAccount = xxxBoolean("ADS_UF_MNS_LOGON_ACCOUNT", objSiteChild.userAccountControl, ADS_UF_MNS_LOGON_ACCOUNT, xUserAccountControl)
			Set xUACSmartcardRequired = xxxBoolean("ADS_UF_SMARTCARD_REQUIRED", objSiteChild.userAccountControl, ADS_UF_SMARTCARD_REQUIRED, xUserAccountControl)
			Set xUACTrustedForDelegation = xxxBoolean("ADS_UF_TRUSTED_FOR_DELEGATION", objSiteChild.userAccountControl, ADS_UF_TRUSTED_FOR_DELEGATION, xUserAccountControl)
			Set xUACNotDelegation = xxxBoolean("ADS_UF_NOT_DELEGATED", objSiteChild.userAccountControl, ADS_UF_NOT_DELEGATED, xUserAccountControl)
			Set xUACUseDesKeyOnly = xxxBoolean("ADS_UF_USE_DES_KEY_ONLY", objSiteChild.userAccountControl, ADS_UF_USE_DES_KEY_ONLY, xUserAccountControl)
			Set xUACDontReqPreauth = xxxBoolean("ADS_UF_DONT_REQ_PREAUTH", objSiteChild.userAccountControl, ADS_UF_DONT_REQ_PREAUTH, xUserAccountControl)
			Set xUACPasswordExpired = xxxBoolean("ADS_UF_PASSWORD_EXPIRED", objSiteChild.userAccountControl, ADS_UF_PASSWORD_EXPIRED, xUserAccountControl)
			Set xUACTrustedToAuthForDelegation = xxxBoolean("ADS_UF_TRUSTED_TO_AUTH_FOR_DELEGATION", objSiteChild.userAccountControl, ADS_UF_TRUSTED_TO_AUTH_FOR_DELEGATION, xUserAccountControl)
			Set xUACPartialSecretsAccount = xxxBoolean("ADS_UF_PARTIAL_SECRETS_ACCOUNT", objSiteChild.userAccountControl, ADS_UF_PARTIAL_SECRETS_ACCOUNT, xUserAccountControl)
			'********************************************************************
		End if

		'********************************************************************
		'* This section reports the Exchange attributes
		Set xEmail = CreateXMLDOMElement("mail", objSiteChild, xRecord)



		Set xMsExchHomeServerName = CreateXMLDOMElement("msExchHomeServerName", objSiteChild, xRecord)
		Set xMsExchMobileMailboxFlags = CreateXMLDOMElement("msExchMobileMailboxFlags", objSiteChild, xRecord)

		Set xMsExchRecipientDisplayType = CreateXMLDOMElement("msExchRecipientDisplayType", objSiteChild, xRecord)
		Set xMsExchRecipientTypeDetails = CreateXMLDOMElement("msExchRecipientTypeDetails", objSiteChild, xRecord)
		Set xLegacyExchangeDN = CreateXMLDOMElement("legacyExchangeDN", objSiteChild, xRecord)

		
		Set xMDBStorageQuota = CreateXMLDOMElement("mDBStorageQuota", objSiteChild, xRecord)
		Set xMDBOverQuotaLimit = CreateXMLDOMElement("mDBOverQuotaLimit", objSiteChild, xRecord)
		Set xMDBUseDefaults = CreateXMLDOMElement("mDBUseDefaults", objSiteChild, xRecord)
		Set xHomeMDB = CreateXMLDOMElement("homeMDB", objSiteChild, xRecord)
		Set xHomeMTA = CreateXMLDOMElement("HomeMTA", objSiteChild, xRecord)

		'********************************************************************


		'********************************************************************
		'* This section reports the Office Communicator 2007 attributes

'		If isEmpty(objSiteChild.msRTCSIP-xxx) = FALSE Then
			Set xMsRTCSIP = xDoc.createElement("msRTCSIP")
			xRecord.appendChild xMsRTCSIP
			Set xMsRTCSIPUserEnabled = CreateXMLDOMElement("msRTCSIP-UserEnabled", objSiteChild, xMsRTCSIP)
			Set xMsRTCSIPFederationEnabled = CreateXMLDOMElement("msRTCSIP-FederationEnabled", objSiteChild, xMsRTCSIP)
			Set xMsRTCSIPInternetAccessEnabled = CreateXMLDOMElement("msRTCSIP-InternetAccessEnabled", objSiteChild, xMsRTCSIP)
			Set xMsRTCSIPOptionFlags = CreateXMLDOMElement("msRTCSIP-OptionFlags", objSiteChild, xMsRTCSIP)
			Set xMsRTCSIPPrimaryHomeServer = CreateXMLDOMElement("msRTCSIP-PrimaryHomeServer", objSiteChild, xMsRTCSIP)
			Set xMsRTCSIPPrimaryUserAddress = CreateXMLDOMElement("msRTCSIP-PrimaryUserAddress", objSiteChild, xMsRTCSIP)
		'********************************************************************
			
		'********************************************************************
		'* PwdLastSet, pwdMuchChange and lastLogonTimestamp
		'* Converts and display the PwdLastSet value as, Password Much Change (True/False) and date
		'* Converts and display the lastLogonTimestamp value if it is not empty 
		If isEmpty(objSiteChild.pwdLastSet) = FALSE Then
			intPwdLastSet = ExtractCommon_Integer8ToInteger(objSiteChild.pwdLastSet)
			sPwdLastSet = ExtractCommon_IntegerToDate(intPwdLastSet)
			intDateDiffCheck = DateDiff("d",sPwdLastSet,Now)
			
			Set xPwdLastSet = xDoc.createElement("PasswordLastSet")
			Set xPwdMuchChange = xDoc.createElement("MustChangePasswordNextLogin")
			Set xLastLogonTimestamp = xDoc.createElement("LastLogonTimestampGMT")
			Set xLockoutTime = xDoc.createElement("LockoutTime")

			If intPwdLastSet = 0 Then
				xPwdLastSet.Text = "0"
				xPwdMuchChange.Text = "TRUE"
			Else
				xPwdLastSet.Text = sPwdLastSet
				xPwdMuchChange.Text = "FALSE"
			End if
			'********************************************************************
			'*lastLogonTimestamp
			If IsEmpty(objSiteChild.lastLogonTimestamp) = FALSE Then
				intLastLogonTimestamp = ExtractCommon_Integer8ToInteger(objSiteChild.lastLogonTimestamp)
				If intLastLogonTimestamp = 0 Then
					xLastLogonTimestamp.Text = "0"
				Else
					xLastLogonTimestamp.Text = ExtractCommon_IntegerToDate(intLastLogonTimestamp)
				End if
			End if

			'********************************************************************
			'*lockoutTime
			If IsEmpty(objSiteChild.lockoutTime) = FALSE Then
				intLockoutTime = ExtractCommon_Integer8ToInteger(objSiteChild.lockoutTime)
				If intLockoutTime = 0 Then
					xLockoutTime.Text = "0"
				Else
					xLockoutTime.Text = ExtractCommon_IntegerToDate(intLockoutTime)
				End if
			End if

		xRecord.appendChild xPwdLastSet
		xRecord.appendChild xPwdMuchChange
		xRecord.appendChild xLastLogonTimestamp
		xRecord.appendChild xLockoutTime
	End if
	'********************************************************************
			
					
		'********************************************************************
		'* memberOfCounter
		If IsEmpty(objSiteChild.memberOf) = FALSE OR objSiteChild.hideDLMembership = TRUE Then
			Set xMembersOf = xDoc.createElement("MembersOf")
			Set xMemberOfCount = xDoc.createElement("MemberOfCount")
				
			If objSiteChild.hideDLMembership = TRUE Then
				xMemberOfCount.Text = "Hidden"
			Else
				colMembersOf = objSiteChild.memberOf
				If IsArray(colMembersOf) = TRUE Then
					memberOfCounter =  0
					For each sGroupMemberOf in colMembersOf
						Set xMembersElement = CreateXMLDOMElementFromString("MemberOf", sGroupMemberOf, xMembersOf)
						memberOfCounter = memberOfCounter + 1
					Next
				Else
					Set xMembersElement = CreateXMLDOMElementFromString("MemberOf", colMembersOf, xMembersOf)
					memberOfCounter =  1
				End if					
				xMemberOfCount.Text = memberOfCounter
			End if
			xRecord.appendChild xMembersOf
			xRecord.appendChild xMemberOfCount
		End if
		'********************************************************************
		

		'********************************************************************
		'* memberCounter
		If IsEmpty(objSiteChild.member) = FALSE OR objSiteChild.hideDLMembership = TRUE Then
			Set xMembers = xDoc.createElement("Members")
			Set xMembersCount = xDoc.createElement("MembersCount")
			
			If objSiteChild.hideDLMembership = TRUE Then
				xMembersCount.Text = "Hidden"
			Else
				colMembers = objSiteChild.member
				If IsArray(colMembers) = TRUE Then
					memberCounter = 0
					For each sGroupMember in colMembers
						Set xMembersElement = CreateXMLDOMElementFromString("Member", sGroupMember, xMembers)
						memberCounter = memberCounter + 1
					Next
				Else
					Set xMembersElement = CreateXMLDOMElementFromString("Member", colMembers, xMembers)
					memberCounter = 1
				End if
				xMembersCount.Text = memberCounter
			End if
			xRecord.appendChild xMembers
			xRecord.appendChild xMembersCount
		End if
		'********************************************************************
			
			
		'********************************************************************
		'* Group
		If isEmpty(objSiteChild.groupType) = FALSE Then
			Set xGroupType = xDoc.createElement("GroupType")
			xRecord.appendChild xGroupType

			Set xGroupTypeValue = CreateXMLDOMElement("GroupTypeValue", objSiteChild.groupType, xGroupType)

			Set xGroupTypeStr = xDoc.createElement("GroupTypeStr")
			Select Case objSiteChild.groupType
				Case ADS_UF_SECURITY_GROUP
					xGroupTypeStr.Text = "Security Group"
				Case else
					xGroupTypeStr.Text = "Distribution Group"
				End Select
			xGroupType.appendChild xGroupTypeStr
			
			Set xGroupCreatedBySystem = xDoc.createElement("CreatedBySystem")
			Select Case objSiteChild.groupType
				Case ADS_UF_SYSTEM
					xGroupCreatedBySystem.Text = "TRUE"
				Case else
					xGroupCreatedBySystem.Text = "FALSE"
				End Select
			xGroupType.appendChild xGroupCreatedBySystem

			Set xGroupScope = xDoc.createElement("GroupScope")
			Select Case objSiteChild.groupType
				Case ADS_UF_SCOPE_GLOBAL
					xGroupScope.Text = "Global"
				Case ADS_UF_SCOPE_LOCAL
					xGroupScope.Text = "Local"
				Case ADS_UF_SCOPE_UNIVERSAL
					xGroupScope.Text = "Universal"
				End Select
			xGroupType.appendChild xGroupScope

		End if
		'********************************************************************

		Next
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

If WScript.Arguments.Count = 3 Then
	strType = objArgs(0)
	strScope = objArgs(1)
	strDN = objArgs(2)

		Select Case strType
		Case "1"
			strType = ADS_OBJ_TYPE_COMPUTER
		Case "2"
			strType = ADS_OBJ_TYPE_GROUP
		Case "3"
			strType = ADS_OBJ_TYPE_USER
		End Select
	
Else
	strType = InputBox("Enter the type object to extract from Active Directory Computer[1], Group[2], or User[3] ","Input Object Type",defaultType)
	  'strType = defaultType
'	If strType = "1" OR strType = "2" OR strType = "3" Then




		Select Case strType
		Case "1"
			strType = ADS_OBJ_TYPE_COMPUTER
		Case "2"
			strType = ADS_OBJ_TYPE_GROUP
		Case "3"
			strType = ADS_OBJ_TYPE_USER
		End Select
	'********************************************************************

	'********************************************************************
	'* InputBox to select mode of operation. Extract data from: Single OU or Recursively 
	strScope = InputBox("Enter the mode of operation. Single OU Only  [1] or additionally query child OU [2]","Input Scope",defaultScope)
	'	strScope = defaultScope
'		If strScope = "1" OR strScope = "2" Then
	'* InputBox to select the DN of the target OU
	strDN = InputBox("Enter the distinguished name of a Site container","Input Site OU",defaultOU)
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









