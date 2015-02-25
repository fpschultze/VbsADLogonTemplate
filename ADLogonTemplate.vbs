'
' =================================================================
'
'   Script Information
'   ------------------
'
'   Filename:       logon.vbs
'   Description:    Active Directory Logon Script
'   Created:        9 Aug 2006
'   Author:         Frank-Peter Schultze
'
' =================================================================
'
'
'-- Need to declare variables
'
Option Explicit
'
'-- Define MAP group prefixes
'
Const MAP_DRIVE_GROUP_PREFIX = "MAP-DRV-"
Const MAP_PRINTER_GROUP_PREFIX = "MAP-PRN-"
'
'-- Global variables
'
Dim blnSilent
Dim strLastErrSource, strLastErrNumber, strLastErrDescription
Dim objNet, objUser
Dim strLogonName, strUserDomain, strUserDN
Dim objMemberOf, arrMemberOf, intGroupCount
Dim strGroup, strGroupDN, objGroup
Dim strNetUse, arrNetUse
Dim strDrv, strUnc
Dim strPrinterPath
Dim i, j, l, u
'
'-- Silent operation/do not show errors?
'
blnSilent = True
'
'-- Bind the user object . . .
'
Set objNet = WScript.CreateObject("Wscript.Network")
strLogonName = UCase(objNet.UserName)
strUserDomain = objNet.UserDomain
strUserDN = GetDN(strUserDomain & "\" & strLogonName)
Set objUser = GetObject("LDAP://" & strUserDN)
'
'-- Load the user's direct and indirect group memberships . . .
'
Set objMemberOf = WScript.CreateObject("Scripting.Dictionary")
LoadGroups strLogonName, objUser
arrMemberOf = objMemberOf.Keys
intGroupCount = objMemberOf.Count
'
'-- Map the user's network drives . . .
'
For i = 0 To intGroupCount - 1
	strGroup = Replace(arrMemberOf(i), strLogonName & "\", "")
	j = Len(MAP_DRIVE_GROUP_PREFIX)
	If (Left(strGroup, j) = MAP_DRIVE_GROUP_PREFIX) Then
		strGroupDN = GetDN(strUserDomain & "\" & strGroup)
		Set objGroup = GetObject("LDAP://" & strGroupDN)
		On Error Resume Next
		strNetUse = Trim(CStr(objGroup.Description))
		If (strNetUse <> "") Then
			arrNetUse = Split(strNetUse, " ")
			l = LBound(arrNetUse)
			u = UBound(arrNetUse)
			If (l <> u) Then
				strDrv = arrNetUse(l)
				strUnc = arrNetUse(u)
				If Not MapDrive(strDrv, strUnc) Then
					TimeBombedErrorBox "MapDrive(" & Chr(34) & strDrv & Chr(34) & ", " _
						& Chr(34) & strUnc & Chr(34) & ") returned an error!"
				End If
			End If
		End If
		On Error GoTo 0
		Set objGroup = Nothing
	End If
Next
'
'-- Map the user's network printers . . .
'
For i = 0 To intGroupCount - 1
	strGroup = Replace(arrMemberOf(i), strLogonName & "\", "")
	j = Len(MAP_PRINTER_GROUP_PREFIX)
	If (Left(strGroup, j) = MAP_PRINTER_GROUP_PREFIX) Then
		strGroupDN = GetDN(strUserDomain & "\" & strGroup)
		Set objGroup = GetObject("LDAP://" & strGroupDN)
		On Error Resume Next
		strPrinterPath = CStr(objGroup.Description)
		If (strPrinterPath <> "") Then
			objNet.AddWindowsPrinterConnection strPrinterPath
			If (Err.Number <> 0) Then
				strLastErrSource = CStr(Err.Source)
				strLastErrNumber = CStr(Err.Number)
				strLastErrDescription = CStr(Err.Description)
				Err.Clear
				TimeBombedErrorBox "AddWindowsPrinterConnection " & Chr(34) _
					& strPrinterPath & Chr(34) & " returned an error!"
			End If
		End If
		On Error GoTo 0
		Set objGroup = Nothing
	End If
Next

Set objMemberOf = Nothing
Set objUser = Nothing
Set objNet = Nothing

'//
'// Procedures
'//

Sub LoadGroups(strLogonName, objADObject)

	'
	' Load the user's direct and indirect group memberships
	'

	Dim colstrGroups
	Dim objGroup
	Dim j

	objMemberOf.CompareMode = vbTextCompare
	colstrGroups = objADObject.memberOf
	If IsEmpty(colstrGroups) Then
		Exit Sub
	End If
	If TypeName(colstrGroups) = "String" Then
		Set objGroup = GetObject("LDAP://" & colstrGroups)
		If Not objMemberOf.Exists(strLogonName & "\" & objGroup.sAMAccountName) Then
			objMemberOf(strLogonName & "\" & objGroup.sAMAccountName) = True
			LoadGroups strLogonName, objGroup
		End If
		Set objGroup = Nothing
		Exit Sub
	End If
	For j = 0 To UBound(colstrGroups)
		Set objGroup = GetObject("LDAP://" & colstrGroups(j))
		If Not objMemberOf.Exists(strLogonName & "\" & objGroup.sAMAccountName) Then
			objMemberOf(strLogonName & "\" & objGroup.sAMAccountName) = True
			LoadGroups strLogonName, objGroup
		End If
	Next
	Set objGroup = Nothing

End Sub

Function MapDrive(strDrive, strShare)

	'
	' Map a network drive
	'

	Const DRIVE_TYPE_NETWORK = 3

	Dim objDrive
	Dim objFSO
	Dim objNet

	MapDrive = False
	strDrive = Left(strDrive, 2)
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set objNet = WScript.CreateObject("WScript.Network")
	On Error Resume Next
	If objFSO.DriveExists(strDrive) Then
		Set objDrive = objFSO.GetDrive(strDrive)
		If (objDrive.DriveType <> DRIVE_TYPE_NETWORK) Then
			strLastErrSource = "MapDrive"
			strLastErrNumber = ""
			strLastErrDescription = "Drive letter already in use"
			On Error GoTo 0
			Exit Function
		End If
		objNet.RemoveNetworkDrive strDrive, True, True
		If (Err.Number <> 0) Then
			strLastErrSource = CStr(Err.Source)
			strLastErrNumber = CStr(Err.Number)
			strLastErrDescription = CStr(Err.Description)
			Err.Clear
			On Error GoTo 0
			Exit Function
		End If
	End If
	objNet.MapNetworkDrive strDrive, strShare
	If (Err.Number <> 0) Then
		strLastErrSource = CStr(Err.Source)
		strLastErrNumber = CStr(Err.Number)
		strLastErrDescription = CStr(Err.Description)
		Err.Clear
	Else
		MapDrive = True
	End If
	On Error GoTo 0
	Set objFSO = Nothing

End Function

Sub TimeBombedErrorBox(strMsg)

	Dim objWsh
	Dim intTimeout
	Dim strTitle
	Dim intButton

	If (blnSilent <> True ) Then
		Set objWsh = WScript.CreateObject("WScript.Shell")
		intTimeout = 30
		strTitle = "Error"
		strMsg = strMsg & vbNewLine & vbNewLine
		strMsg = strMsg & "Error source: " & strLastErrSource & vbNewLine
		strMsg = strMsg & "Error number: " & strLastErrNumber & vbNewLine
		strMsg = strMsg & "Error description: " & strLastErrDescription
		intButton = objWsh.Popup(strMsg, intTimeout, strTitle, vbExclamation)
		Set objWsh = Nothing
	End If

End Sub

Function GetDN(strNTName)

	'
	' Translate NT4 name to its DN
	'

	Const ADS_NAME_INITTYPE_GC = 3
	Const ADS_NAME_TYPE_NT4 = 3
	Const ADS_NAME_TYPE_1779 = 1

	Dim objTrans

	Set objTrans = WScript.CreateObject("NameTranslate")
	objTrans.Init ADS_NAME_INITTYPE_GC, ""
	objTrans.Set ADS_NAME_TYPE_NT4, strNTName
	GetDN = objTrans.Get(ADS_NAME_TYPE_1779)
	Set objTrans = Nothing

End Function
