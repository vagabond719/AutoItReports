#AutoIt3Wrapper_Res_File_Add=capture.jpg, rt_rcdata, TEST_JPG_1
#region Includes
#include <Excel.au3>
#include <Array.au3>
#include <Date.au3>
#include <OutlookEx.au3>
#include <Parse.au3>
#include <GUIConstantsEx.au3>
#endregion Includes
#region Variable declaration
$fwcday = _DateToDayOfWeekISO(@YEAR, @MON, @MDAY)
$fwcurrent = @YEAR & "/" & @MON & "/" & @MDAY
$fwcday -= 1
$fwcday = "-" & $fwcday
$fwcurrentstart = _DateAdd('d', $fwcday, $fwcurrent)
$fwfday = _DateToDayOfWeekISO(@YEAR, 01, 01)
$fwfirst = @YEAR & "/01/01"
$fwfday -= 1
$fwfday = "-" & $fwfday
$fwweekstart = _DateAdd('d', $fwfday, $fwfirst)
$fwweekend = _DateAdd('d', 6, $fwweekstart)
$fwweekcount = ""
$date = @YEAR & @MON & @MDAY
$fday = _DateToDayOfWeekISO(@YEAR, 01, 01)
$first = @YEAR & "/01/01"
$fday -= 1
$fday = "-" & $fday
$weekstart = _DateAdd('d', $fday, $first)
$weekend = _DateAdd('d', 6, $weekstart)
#endregion Variable declaration
#region Calculate current fiscal week
For $fwcounter = 1 To 52
	$fwweekcount += 1
	$fwdiff = _DateDiff("d", $fwcurrent, $fwweekend)
	If $fwdiff >= 1 And $fwdiff <= 7 Then
		ExitLoop
	EndIf
	$fwweekend = _DateAdd('w', 1, $fwweekend)
Next
#endregion Calculate current fiscal week
#region Holiday Array
Dim $holiday[21][2]
$holiday[1][0] = "New Year's Day"
$holiday[2][0] = "Martin Luther King"
$holiday[3][0] = "Presidents' Day"
$holiday[4][0] = "Memorial Day"
$holiday[5][0] = "Independence Day"
$holiday[6][0] = "Labor Day"
$holiday[7][0] = "Columbus Day"
$holiday[8][0] = "Veterans Day"
$holiday[9][0] = "Thanksgiving Day"
$holiday[10][0] = "Christmas"
$holiday[11][0] = "Last New Year's Day"
$holiday[12][0] = "Last Martin Luther King"
$holiday[13][0] = "Last Presidents' Day"
$holiday[14][0] = "Last Memorial Day"
$holiday[15][0] = "Last Independence Day"
$holiday[16][0] = "Last Labor Day"
$holiday[17][0] = "Last Columbus Day"
$holiday[18][0] = "Last Veterans Day"
$holiday[19][0] = "Last Thanksgiving Day"
$holiday[20][0] = "Last Christmas"

$mlk = _DateToDayOfWeekISO(@YEAR, "01", "15")
If $mlk <> 1 Then
	$mlk = @YEAR & "/01/" & 15 + (8 - $mlk)
ElseIf $mlk = 1 Then
	$mlk = @YEAR & "/01/15"
EndIf

$lmlk = _DateToDayOfWeekISO((@YEAR - 1), "01", "15")
If $lmlk <> 1 Then
	$lmlk = (@YEAR - 1) & "/01/" & 15 + (8 - $lmlk)
ElseIf $lmlk = 1 Then
	$lmlk = (@YEAR - 1) & "/01/15"
EndIf

$pres = _DateToDayOfWeekISO(@YEAR, "02", "15")
If $pres <> 1 Then
	$pres = @YEAR & "/02/" & 15 + (8 - $pres)
ElseIf $pres = 1 Then
	$pres = @YEAR & "02/15"
EndIf

$lpres = _DateToDayOfWeekISO((@YEAR - 1), "02", "15")
If $lpres <> 1 Then
	$lpres = (@YEAR - 1) & "/02/" & 15 + (8 - $lpres)
ElseIf $lpres = 1 Then
	$lpres = (@YEAR - 1) & "02/15"
EndIf

$memor = _DateToDayOfWeekISO(@YEAR, "05", "31")
If $memor <> 1 Then
	$memor = @YEAR & "/05/" & 32 - $memor
ElseIf $memor = 1 Then
	$memor = @YEAR & "/05/31"
EndIf

$lmemor = _DateToDayOfWeekISO((@YEAR - 1), "05", "31")
If $lmemor <> 1 Then
	$lmemor = (@YEAR - 1) & "/05/" & 32 - $lmemor
ElseIf $lmemor = 1 Then
	$lmemor = (@YEAR - 1) & "/05/31"
EndIf

$lab = _DateToDayOfWeekISO(@YEAR, "09", "01")
If $lab <> 1 Then
	$lab = @YEAR & "/09/" & 1 + (8 - $lab)
ElseIf $lab = 1 Then
	$lab = @YEAR & "/09/01"
EndIf

$llab = _DateToDayOfWeekISO((@YEAR - 1), "09", "01")
If $llab <> 1 Then
	$llab = (@YEAR - 1) & "/09/" & 1 + (8 - $llab)
ElseIf $llab = 1 Then
	$llab = (@YEAR - 1) & "/09/01"
EndIf

$columbus = _DateToDayOfWeekISO(@YEAR, "10", "08")
If $columbus <> 1 Then
	$columbus = @YEAR & "/10/" & 8 + (8 - $columbus)
ElseIf $columbus = 1 Then
	$columbus = @YEAR & "/10/08"
EndIf

$lcolumbus = _DateToDayOfWeekISO((@YEAR - 1), "10", "08")
If $lcolumbus <> 1 Then
	$lcolumbus = (@YEAR - 1) & "/10/" & 8 + (8 - $lcolumbus)
ElseIf $lcolumbus = 1 Then
	$lcolumbus = (@YEAR - 1) & "/10/08"
EndIf

$thank = _DateToDayOfWeekISO(@YEAR, "11", "22")
If $thank <> 4 Then
	$thank = @YEAR & "/11/" & 22 + (8 - $thank)
ElseIf $thank = 4 Then
	$thank = @YEAR & "/11/22"
EndIf

$lthank = _DateToDayOfWeekISO((@YEAR - 1), "11", "22")
If $lthank <> 4 Then
	$lthank = (@YEAR - 1) & "/11/" & 22 + (8 - $lthank)
ElseIf $lthank = 4 Then
	$lthank = (@YEAR - 1) & "/11/22"
EndIf

$holiday[1][1] = @YEAR & "/01/01"
$holiday[2][1] = $mlk
$holiday[3][1] = $pres
$holiday[4][1] = $memor
$holiday[5][1] = @YEAR & "/07/04"
$holiday[6][1] = $lab
$holiday[7][1] = $columbus
$holiday[8][1] = @YEAR & "/11/11"
$holiday[9][1] = $thank
$holiday[10][1] = @YEAR & "/12/25"
$holiday[11][1] = (@YEAR - 1) & "/01/01"
$holiday[12][1] = $lmlk
$holiday[13][1] = $lpres
$holiday[14][1] = $lmemor
$holiday[15][1] = (@YEAR - 1) & "/07/04"
$holiday[16][1] = $llab
$holiday[17][1] = $lcolumbus
$holiday[18][1] = (@YEAR - 1) & "/11/11"
$holiday[19][1] = $lthank
$holiday[20][1] = (@YEAR - 1) & "/12/25"

$holiday[0][0] = UBound($holiday) - 1
#endregion Holiday Array
#region Download task.xls
If FileExists(@ScriptDir & "\u_service_request.xls") Then
	FileDelete(@ScriptDir & "\u_service_request.xls")
EndIf
If FileExists(@ScriptDir & "\incident.xls") Then
	FileDelete(@ScriptDir & "\incident.xls")
EndIf

$command = Run(@ScriptDir & "\servicenow.exe", @ScriptDir,@SW_HIDE)

While ProcessExists($command)
   sleep(100)
WEnd
#endregion Download task.xls
#region File setup !!Calls function Main!!

$folderdate = @ScriptDir & "\" & $date & "_" & "daily\"
DirCreate($folderdate)
FileCopy(@ScriptDir & "\Template\TSG.XLS", @ScriptDir, 1)
FileCopy(@ScriptDir & "\u_service_request.xls", $folderdate, 1)
FileCopy(@ScriptDir & "\incident.xls", $folderdate, 1)
FileMove($folderdate & "u_service_request.xls", $folderdate & "u_service_request_raw.xls")
FileMove($folderdate & "incident.xls", $folderdate & "incident_raw.xls")

Main(1)
Main(2)
$subject = "**Daily SLA Report**"
$to = "502117737"
$cc = ""
$bcc = "502183923"
$body = ""
$attachment = (@ScriptDir & "\" & "tsg.xls")
;~ Global $oOL = _OL_Open()
;~ _OL_Wrapper_SendMail($oOL, $to, $cc, $bcc, $subject, $body, $attachment, $olFormatHTML, $olImportanceNormal)
FileCopy(@ScriptDir & "\TSG.XLS", @ScriptDir & "\TSGNEW.XLS", 1)
FileMove(@ScriptDir & "\TSG.XLS", $folderdate, 9)
FileMove(@ScriptDir & "\u_service_request.xls", $folderdate & "u_service_request_processed.xls")
FileMove(@ScriptDir & "\incident.xls", $folderdate & "incident_processed.xls")
Run(@ScriptDir & "\SLAaudit v2.5.exe", @ScriptDir,@SW_HIDE)

#endregion File setup !!Calls function Main!!
Func Main($srorinc)
	;------------------------------------Set file name based on passed variable
	If $srorinc = 1 Then
		Local $sFilePath1 = @ScriptDir & "\u_service_request.xls"
		$table = "ServiceRequests"
		$table2 = "Bus Days minus Exception"
	ElseIf $srorinc = 2 Then
		Local $sFilePath1 = @ScriptDir & "\incident.xls"
		$table = "Incidents"
		$table2 = "Bus Dur minus Exception"
	EndIf
	;------------------------------------Confirm file exists check
	If Not FileExists($sFilePath1) Then
		MsgBox(16, '', 'Does NOT exists')
		Exit
	EndIf
	;------------------------------------Open Excel and insert rows
	Local $oExcel = _ExcelBookOpen($sFilePath1, 0)
	_ExcelSheetActivate($oExcel, "Sheet1")
	_ExcelColumnInsert($oExcel, 17, 2)
	_ExcelWriteCell($oExcel, "RAW", 1, 17)
	_ExcelWriteCell($oExcel, "ADJ", 1, 18)
	_ExcelWriteCell($oExcel, "RAW_SLA_Made", 1, 19)
	_ExcelWriteCell($oExcel, "ADJ_SLA_Made", 1, 20)
	_ExcelBookSave($oExcel)
	_ExcelBookClose($oExcel, 0, 0)
	;------------------------------------Build Excel connections
	$conn = ObjCreate("ADODB.Connection")
	$DSN = ("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & $sFilePath1 & ";readOnly=false;Nullable=true")
	$conn.Open($DSN)

	$conn2 = ObjCreate("ADODB.Connection")
	$file = @ScriptDir & "\tsg.xls"
	$DSN2 = ("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & $file & ";readOnly=false;Nullable=true")
	$conn2.Open($DSN2)
	#region Short description cleanup
	$query2 = 'update [Page 1$] set "Short description"=' & "'privilaged' where" & ' "Short description" like ' & "'%Automated HPA%' and " & '"Assignment group"' & _
			"='L2 TSG Account Admin Support'"
	$conn.execute($query2)
	$query2 = 'update [Page 1$] set "Short description"=' & "'e2k' where" & '"Assignment group"' & "='L2 TSG E2K Support Team'"
	$conn.execute($query2)
	$query2 = 'update [Page 1$] set "Short description"' & "='unix' where" & ' "Short description" like ' & "'%unix%' or" & _ ; & ' "Short description" like ' & "'%hpa%' or"
			"" & ' "Short description" like ' & "'%sudo%' or" & ' "Short description" like ' & "'%ftp%' or" & ' "Short description" like ' & "'%afs%' or" & ' "Short description" like ' & "'%tns%' or" & ' "Short description" like ' & "'linux%' and " & '"' & _
			"Assignment group" & '"' & "='L2 TSG Account Admin Support'"
	$conn.execute($query2)
	$query2 = 'update [Page 1$] set "Short description"' & "='vpn' where" & ' "Short description" like ' & "'%vpn%' or" & ' "Short description" like ' & "'%3rd%' or" & _
			"" & ' "Short description" like ' & "'%stag%' or" & ' "Short description" like ' & "'%token%' or" & ' "Short description" like ' & "'%pin%' or" & ' "Short description" like ' & "'%secur%' or" & ' "Short description" like ' & "'deprovision%' or" & _
			"" & ' "Short description" like ' & "'keyfob%' and " & '"Assignment group"' & "='L2 TSG Account Admin Support'"
	$conn.execute($query2)
	$query2 = 'update [Page 1$] set "Short description"' & "='group' where" & ' "Short description" like ' & "'%GRP%' or" & ' "Short description" like ' & "'%Group%' and " & '"' & _
			"Assignment group" & '"' & "='L2 TSG Account Admin Support'"
	$conn.execute($query2)
	$query2 = 'update [Page 1$] set "Short description"' & "='share' where" & ' "Short description" like ' & "'%share%' or" & ' "Short description" like ' & "'%shr%' or" & _
			"" & ' "Short description" like ' & "'%access%' or" & ' "Short description" like ' & "'%folder%' and " & '"' & "Assignment group" & '"' & "='L2 TSG Account Admin Support'"
	$conn.execute($query2)
	$query2 = 'update [Page 1$] set "Short description"' & "='password reset' where" & ' "Short description" like ' & "'%password%' and " & '"' & "Assignment group" & '"' & _
			"='L2 TSG Account Admin Support'"
	$conn.execute($query2)
	$query2 = 'update [Page 1$] set "Short description"' & "='nuclear nt/email' where" & ' "Short description" like ' & "'%nuclear%' and " & '"' & "Assignment group" & '"' & _
			"='L2 TSG Account Admin Support'"
	$conn.execute($query2)
	$query2 = 'update [Page 1$] set "Short description"' & "='modifyntacct' where" & ' "Short description" like ' & "'%ModifyNTAcct%' or" & ' "Short description" like ' & "'%ShopFloor%' and " & _
			'"' & "Assignment group" & '"' & "='L2 TSG Account Admin Support'"
	$conn.execute($query2)
	$query2 = 'update [Page 1$] set "Short description"' & "='disabled domain' where" & ' "Short description" like ' & "'%able%' or" & ' "Short description" like ' & "'%login%' or" & _
			"" & ' "Short description" like ' & "'%windows%' or" & ' "Short description" like ' & "'%domain%' or" & ' "Short description" like ' & "'%lock%' and " & '"' & "Assignment group" & '"' & "='L2 TSG Account Admin Support'"
	$conn.execute($query2)
	$query2 = 'update [Page 1$] set "Short description"=' & "'other' where" & '"Short description" not in' & "('privilaged','unix','vpn','group','share','password reset','nuclear nt/email','modifyntacct','disabled domain') and " & '"Assignment group"' & _
			"='L2 TSG Account Admin Support'"
	$conn.execute($query2)
	#endregion Short description cleanup
	#region Region code cleanup
	$query2 = "update [Page 1$] set location='AM' where location like 'am%'"
	$conn.execute($query2)
	$query2 = "update [Page 1$] set location='EM' where location like 'em%'"
	$conn.execute($query2)
	$query2 = "update [Page 1$] set location='AS' where location like 'as%'"
	$conn.execute($query2)
	$query2 = "update [Page 1$] set location='Missing Region' where location not in ('AM','EM','AS')"
	$conn.execute($query2)
	#endregion Region code cleanup
	#region Business cleanup
	$query2 = "update [Page 1$] set " & '"' & "Business segment" & '"' & "='energy' where " & '"' & "Business segment" & '"' & " like '%energy%'"
	$conn.execute($query2)
	$query2 = "update [Page 1$] set " & '"' & "Business segment" & '"' & "='energy' where " & '"' & "Business segment" & '"' & " like '%power & water%'"
	$conn.execute($query2)
	$query2 = "update [Page 1$] set " & '"' & "Business segment" & '"' & "='aviation' where " & '"' & "Business segment" & '"' & " like '%aviation%'"
	$conn.execute($query2)
	$query2 = "update [Page 1$] set " & '"' & "Business segment" & '"' & "='aviation' where " & '"' & "Business segment" & '"' & " like '%transportation%'"
	$conn.execute($query2)
	$query2 = "update [Page 1$] set " & '"' & "Business segment" & '"' & "='aviation' where " & '"' & "Business segment" & '"' & " like '%Initiatives%'"
	$conn.execute($query2)
	$query2 = "update [Page 1$] set " & '"' & "Business segment" & '"' & "='ge infrastructure' where " & '"' & "Business segment" & '"' & " like '%Infrastructure%'"
	$conn.execute($query2)
	$query2 = "update [Page 1$] set " & '"' & "Business segment" & '"' & "='energy' where " & '"' & "Business segment" & '"' & " like '%oil & gas%'"
	$conn.execute($query2)
	$query2 = "update [Page 1$] set " & '"' & "Business segment" & '"' & "='aviation' where " & '"' & "Business segment" & '"' & " not in " & _
			"('aviation','energy','ge infrastructure')"
	$conn.execute($query2)
	#endregion Business cleanup
	;------------------------------------Calculate SLA
	$query = "Select * FROM [Page 1$]"
	$aArray = _Parse($conn, $query)
	;_ArrayDisplay($aArray, "Array using Default Parameters")
	;------------------------------------Define needed values to step through calculations
	$arraycount = UBound($aArray)
	ReDim $aArray[$arraycount][25]
	$arraycount -= 1
	$aArray[0][1] += 5

	;------------------------------------Main process loop
	For $loop = 1 To $arraycount
		#region Weekend Calculator
		$start = $aArray[$loop][1] ;Open date of ticket
		$end = $aArray[$loop][2] ;Close date of ticket
		$daycounter = "" ;Tracks the number of weekend days
		$holdaycounter = "" ;Tracks the number of holiday days
		$adjustedstart = "" ;Time passed on start date (weekend)
		$adjustedend = "" ;Time passed on End date (weekend)
		$holadjustedstart = "" ;Time passed on start date (holiday)
		$holadjustedend = "" ;Time passed on End date (holiday)
		;$datecatch = ""
		$differencestart = ((((StringMid($start, 9, 2)) * 3600) + (StringMid($start, 11, 2)) * 60) + StringMid($start, 13, 2)) ;Convert start time difference to sec
		$differneceend = 86400 - ((((StringMid($end, 9, 2)) * 3600) + (StringMid($end, 11, 2)) * 60) + StringMid($end, 13, 2)) ;Convert end time difference to sec
		$holdifferencestart = 86400 - ((((StringMid($start, 9, 2)) * 3600) + (StringMid($start, 11, 2)) * 60) + StringMid($start, 13, 2)) ;Convert start time difference to sec
		$holdifferneceend = ((((StringMid($end, 9, 2)) * 3600) + (StringMid($end, 11, 2)) * 60) + StringMid($end, 13, 2)) ;Convert end time difference to sec
		$dayofstart = _DateToDayOfWeekISO(StringMid($start, 1, 4), StringMid($start, 5, 2), StringMid($start, 7, 2)) ;Check what day of the week the ticket was opened
		$dayofstart2 = _DateToDayOfWeekISO(StringMid($start, 1, 4), StringMid($start, 5, 2), StringMid($start, 7, 2)) ;Counter used when calculating the weekend days
		$dayofend = _DateToDayOfWeekISO(StringMid($end, 1, 4), StringMid($end, 5, 2), StringMid($end, 7, 2)) ;Check what day of the week the ticket was closeed
		$start = StringMid($start, 1, 4) & "/" & StringMid($start, 5, 2) & "/" & StringMid($start, 7, 2) & " " & _
				StringMid($start, 9, 2) & ":" & StringMid($start, 11, 2) & "" & ":" & StringMid($start, 13, 2) ;Convert format of date/time
		$end = StringMid($end, 1, 4) & "/" & StringMid($end, 5, 2) & "/" & StringMid($end, 7, 2) & " " & StringMid($end, 9, 2) & ":" & StringMid($end, 11, 2) & _
				"" & ":" & StringMid($end, 13, 2) ;Convert format of date/time
		$startcalc = StringMid($start, 1, 10) & " 00:00:00" ;Set start time to the begining of the day
		$endcalc = StringMid($end, 1, 10) & " 23:59:59" ;Set end time to the end of the day
		$diffday = _DateDiff('d', $startcalc, $endcalc) ;Determine the number of days between opened and closed
		;------------------------------------Ammendment to support closure code
		$aArray[$loop][23] = $aArray[$loop][14]
		If $aArray[$loop][23] = "" Then
			$aArray[$loop][23] = "Exception"
		ElseIf $aArray[$loop][23] = "Closed by automation" Then
			$aArray[$loop][23] = "Exception"
		ElseIf $aArray[$loop][23] = "Request Cancelled" Then
			$aArray[$loop][23] = "Exception"
		ElseIf $aArray[$loop][23] = "User closed - Issue Disappeared" Then
			$aArray[$loop][23] = "Exception"
		ElseIf $aArray[$loop][23] = "User closed - User Resolved" Then
			$aArray[$loop][23] = "Exception"
		ElseIf $aArray[$loop][23] = "User closed - Resolved by the support team" Then
			$aArray[$loop][23] = "Exception"
		ElseIf $aArray[$loop][4] = "Cancelled" Then
			$aArray[$loop][23] = "Exception"
		EndIf
		If $aArray[$loop][5] = "WILLIAMS, KENNETH (501942315)" Then
			If $aArray[$loop][5] = "Request For Information" Or $aArray[$loop][5] = "Completed Request" Then
				$aArray[$loop][23] = "Exception"
			EndIf
		EndIf
		;------------------------------------
		$aArray[$loop][14] = $diffday ;Write date difference to array
		$diffsec = _DateDiff('s', $start, $end) ;Determine the number of seconds between opened and closed
		;------------------------------------Count the number of weekend days within date range
		;$dayofstart is a representation of the day of the week
		;$daycounter is the value for counting the weekend days
		For $y = 0 To $diffday
			If $dayofstart2 = 7 Then
				$daycounter += 1
				$dayofstart2 = 1
			ElseIf $dayofstart2 = 6 Then
				$daycounter += 1
				$dayofstart2 += 1
			Else
				$dayofstart2 += 1
			EndIf
		Next
		If $dayofstart > 5 Then
			$adjustedstart = $differencestart
		EndIf
		If $dayofend > 5 Then
			$adjustedend = $differneceend
		EndIf
		#endregion Weekend Calculator
		#region Calculate For Holidays
		$hstart = $start ;Open date of ticket
		$hend = $end ;Close date of ticket
		$hdaycounter = 0 ;Tracks the number of weekend days
		$hadjustedstart = "" ;Time passed on start date
		$hadjustedend = "" ;Time passed on End date
		$hdifferencestart = 86400 - ((((StringMid($hstart, 9, 2)) * 3600) + (StringMid($hstart, 11, 2)) * 60) + StringMid($hstart, 13, 2)) ;Convert start time difference to sec
		$hdifferneceend = ((((StringMid($hend, 9, 2)) * 3600) + (StringMid($hend, 11, 2)) * 60) + StringMid($hend, 13, 2)) ;Convert end time difference to sec
		$hstart = StringMid($hstart, 1, 4) & "/" & StringMid($hstart, 5, 2) & "/" & StringMid($hstart, 7, 2) & " " & _
				StringMid($hstart, 9, 2) & ":" & StringMid($hstart, 11, 2) & "" & ":" & StringMid($hstart, 13, 2) ;Convert format of date/time
		$hend = StringMid($hend, 1, 4) & "/" & StringMid($hend, 5, 2) & "/" & StringMid($hend, 7, 2) & " " & StringMid($hend, 9, 2) & ":" & StringMid($hend, 11, 2) & _
				"" & ":" & StringMid($hend, 13, 2) ;Convert format of date/time
		$hstartcalc = StringMid($hstart, 1, 10) & " 00:00:00" ;Set start time to the begining of the day
		$hendcalc = StringMid($hend, 1, 10) & " 23:59:59" ;Set end time to the end of the day
		$hdiffday = _DateDiff('d', $hstartcalc, $hendcalc) ;Determine the number of days between opened and closed
		$hdiffsec = _DateDiff('s', $hstart, $hend)

		$startval = $hstart
		$endval = StringLeft($hend, 10)
		For $q = 1 To $hdiffday + 1
			For $w = 1 To $holiday[0][0]
				If StringLeft($startval, 10) = $holiday[$w][1] And $startval = $hstart Then
					$hadjustedstart = $hdifferencestart
				ElseIf StringLeft($startval, 10) = $holiday[$w][1] And $startval <> $hstart And StringLeft($startval, 10) <> $endval Then
					$hdaycounter += 1
				ElseIf StringLeft($startval, 10) = $holiday[$w][1] And StringLeft($startval, 10) = $endval Then
					$hadjustedend = $hdifferneceend
				EndIf
			Next
			$startval = _DateAdd('D', "1", $startval)
		Next

		$hdaycounter = ($hdaycounter * 86400)
		$hadjustedtime = $hdaycounter + $hadjustedstart + $hadjustedend
		;------------------------------------
		$aArray[$loop][15] = $daycounter
		$daycounter = ($daycounter * 86400)
		$aArray[$loop][24] = $diffsec & " - " & $daycounter & " + " & $adjustedstart & " + " & $adjustedend & " - " & $hadjustedtime
		$adjustedtime = $diffsec - $daycounter + $adjustedstart + $adjustedend - $hadjustedtime
		$aArray[$loop][16] = $adjustedtime
		$aArray[$loop][17] = $adjustedtime - $aArray[$loop][3]
		;------------------------------------Pull SLA weight and convert to seconds
		$sla = StringMid($aArray[$loop][6], 1, 1)
		If $sla = 0 Then
			$sla = 7200
		ElseIf $sla = 1 Then
			$sla = 14400
		ElseIf $sla = 2 Then
			$sla = 86400
		ElseIf $sla = 3 Then
			$sla = 259200
		ElseIf $sla = 4 Then
			$sla = 432000
		EndIf
		;------------------------------------Set missed/made status for RAW numbers
		If $sla >= $aArray[$loop][16] Then
			If $aArray[$loop][2] = "" Then
				$aArray[$loop][19] = "Open"
				$rawmissed = "Open"
				$rawmade = "Open"
				$rawstat = "Open"
			ElseIf $aArray[$loop][2] <> "" Then
				$aArray[$loop][18] = "1"
				$rawmissed = 0
				$rawmade = 1
				$rawstat = "Made"
			EndIf
		ElseIf $sla < $aArray[$loop][16] Then
			$aArray[$loop][18] = "0"
			$rawmissed = 1
			$rawmade = 0
			$rawstat = "Missed"
		EndIf
		;------------------------------------Set missed/made status for Adjusted numbers
		If $sla >= $aArray[$loop][17] Then
			If $aArray[$loop][2] = "" Then
				$aArray[$loop][19] = "Open"
				$adjmissed = "Open"
				$adjmade = "Open"
				$adjstat = "Open"
			ElseIf $aArray[$loop][2] <> "" Then
				$aArray[$loop][19] = "1"
				$adjmissed = 0
				$adjmade = 1
				$adjstat = "Made"
			EndIf
		ElseIf $sla < $aArray[$loop][17] Then
			If $aArray[$loop][23] = "Exception" Then ;Addition to account for closure code
				$aArray[$loop][19] = "1"
				$adjmissed = 0
				$adjmade = 1
				$adjstat = "Made"
			Else
				$aArray[$loop][19] = "0"
				$adjmissed = 1
				$adjmade = 0
				$adjstat = "Missed"
			EndIf
		EndIf
		#endregion Calculate For Holidays
		;------------------------------------
		$bdayminexc = $aArray[$loop][17] / 86400 ;Business day minus exception
		$goal = StringMid($aArray[$loop][6], 1, 1) ;SLA goal of ticket in days
		;------------------------------------Convert Open/Closed date format
		If $aArray[$loop][2] <> "" Then
			$day = StringMid($aArray[$loop][2], 5, 2) & "/" & StringMid($aArray[$loop][2], 7, 2) & "/" & StringMid($aArray[$loop][2], 1, 4)
		ElseIf $aArray[$loop][2] = "" Then
			$day = "open"
		EndIf
		$aArray[$loop][1] = StringMid($aArray[$loop][1], 1, 4) & "/" & StringMid($aArray[$loop][1], 5, 2) & "/" & StringMid($aArray[$loop][1], 7, 2) & " " & _
				StringMid($aArray[$loop][1], 9, 2) & ":" & StringMid($aArray[$loop][1], 11, 2) & "" & ":" & StringMid($aArray[$loop][1], 13, 2)
		If $aArray[$loop][2] <> "" Then
			$aArray[$loop][2] = StringMid($aArray[$loop][2], 1, 4) & "/" & StringMid($aArray[$loop][2], 5, 2) & "/" & StringMid($aArray[$loop][2], 7, 2) & " " & _
					StringMid($aArray[$loop][2], 9, 2) & ":" & StringMid($aArray[$loop][2], 11, 2) & "" & ":" & StringMid($aArray[$loop][2], 13, 2)
		EndIf
		;------------------------------------Calculate the fiscal week the ticket was opened
		$week = ""
		$weekend2 = $weekend
		$opened = StringMid($aArray[$loop][1], 1, 10)
		For $counter = 1 To 52
			$week += 1
			$diff = _DateDiff("d", $opened, $weekend2)
			If $diff >= 1 And $diff <= 7 Then
				$aArray[$loop][20] = $week
				ExitLoop
			EndIf
			$weekend2 = _DateAdd('w', 1, $weekend2)
		Next
		;------------------------------------Calculate the fiscal week the ticket was closed
		$week = ""
		$weekend2 = $weekend
		$opened = StringMid($aArray[$loop][2], 1, 10)
		For $counter = 1 To 52
			$week += 1
			$diff = _DateDiff("d", $opened, $weekend2)
			If $diff >= 1 And $diff <= 7 Then
				$aArray[$loop][21] = $week
				ExitLoop
			EndIf
			$weekend2 = _DateAdd('w', 1, $weekend2)
		Next
		;------------------------------------Calculate the task to time (Open to Close no adjustments)
		$task2time = ""
		If $aArray[$loop][2] <> "" Then
			$task2time = _DateDiff("s", $aArray[$loop][1], $aArray[$loop][2])
			$task2time = (($task2time / 24) / 60) / 60
			$aArray[$loop][22] = $task2time
		EndIf
		;------------------------------------Change all reopen statuses to True
		If $aArray[$loop][13] <> "" Then
			$aArray[$loop][13] = "True"
		EndIf
		;------------------------------------Write array to the TSG.XLS sheet
		$fweek = ""
		$query2 = 'INSERT INTO [' & $table & '$] (ID,Opened,Closed,"Exception time",TasktoTime,"Value",Name,Priority,"Business duration",OPENEDFW,CLOSEDFW,' & _
				'"Short description","Business Segment",Location,Sub,"Assignment group","' & $table2 & '","SLA Goal","Adj SLA Made","Adj SLA Missed",Day,' & _
				'"ADJ MADE/MISSED","RAW MADE","RAW MISSED","RAW MADE/MISSED",Reopened,"Close Code") VALUES ' & _
				"('" & $aArray[$loop][0] & "','" & $aArray[$loop][1] & "','" & $aArray[$loop][2] & "','" & $aArray[$loop][3] & "','" & $aArray[$loop][22] & _
				"','" & $aArray[$loop][4] & "','" & $aArray[$loop][5] & "','" & $aArray[$loop][6] & "','" & $aArray[$loop][16] & "','" & $aArray[$loop][20] & "','" & _
				$aArray[$loop][21] & "','" & $aArray[$loop][8] & "','" & $aArray[$loop][9] & "','" & $aArray[$loop][10] & "','" & $aArray[$loop][11] & "','" & _
				$aArray[$loop][12] & "','" & $bdayminexc & "','" & $goal & "','" & $adjmade & "','" & $adjmissed & "','" & $day & "','" & $adjstat & "','" & _
				$rawmade & "','" & $rawmissed & "','" & $rawstat & "','" & $aArray[$loop][13] & "','" & $aArray[$loop][23] & "');"
		$conn2.execute($query2)
	Next

	;_ArrayDisplay($aArray, "Array using Default Parameters")

	$conn.Close
	$conn2.Close
EndFunc   ;==>Main
