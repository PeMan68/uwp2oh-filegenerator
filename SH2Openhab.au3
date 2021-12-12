; Script generates files for openhab2 and modbus-binding v2
; Parses a configuration-file created by SH tool v 7.x
; The file shall be exported from Reports|Modbus report and saved as .csv
;

#include <File.au3>
#include <Array.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <String.au3>
Global $mIP = "192.168.1.198"
Global $mPort = 502
Global $mID = 1

$filename = Choose_file()
;_InitFiles_v1()
;_ModbusCfgFile(VisaFil($filename))
_InitFiles_v2()
_ModbusThingsFile(VisaFil($filename))
_FormatItemFile()
_CloseFiles()

; Anteckningar

; Rules ska skrivas för varje sekvens som triggas från homekit. Så det skapas en "puls"
;~     rule "Puls Sekvens Kök normal"
;~     when
;~         Item Fx_Kk__Sekvens_S2_Normal_Status changed from OFF to ON
;~     then
;~                  sendCommand(Fx_Kk__Sekvens_S2_Normal_Status,OFF)
;~     end

Func _InitFiles_v1()
	Global $rootfolder = "openhabroot\"
	Local $h = FileOpen($rootfolder & "services\modbus.cfg", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileWriteLine($h, "# Modbus to Smart-House")
	FileClose($h)
	Global $hFileCfg = FileOpen($rootfolder & "services\modbus.cfg", $FO_APPEND + $FO_UTF8_NOBOM)
	;
	Local $h = FileOpen($rootfolder & "items\smart-house.items", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileClose($h)
	Global $hFileItems = FileOpen($rootfolder & "items\smart-house.items", $FO_APPEND + $FO_UTF8_NOBOM)
	;
	Local $h = FileOpen($rootfolder & "rules\smart-house-sequences.rules-temp", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileClose($h)
	Global $hFileRules = FileOpen($rootfolder & "rules\smart-house-sequences.rules-temp", $FO_APPEND + $FO_UTF8_NOBOM)
	;
	Local $h = FileOpen($rootfolder & "items\smart-house-groups.items", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileWriteLine($h, "// Groups")
	FileClose($h)
	Global $hFileItemsGroups = FileOpen($rootfolder & "items\smart-house-groups.items", $FO_APPEND + $FO_UTF8_NOBOM)
	;
	Local $h = FileOpen($rootfolder & "transform\dimmertransform.js", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileWriteLine($h, "(function (inputData) {")
	FileWriteLine($h, "	   return inputData;")
	FileWriteLine($h, "})(input)")
	FileClose($h)
	;
	Local $h = FileOpen($rootfolder & "transform\divide10.js", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileWriteLine($h, "(function (inputData) {")
	FileWriteLine($h, "	   return parseFloat(inputData) / 10;")
	FileWriteLine($h, "})(input)")
	FileClose($h)
	;
	Local $h = FileOpen($rootfolder & "transform\multiply10.js", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileWriteLine($h, "(function (inputData) {")
	FileWriteLine($h, "	   return Math.round(parseFloat(inputData, 10) * 10);")
	FileWriteLine($h, "})(input)")
	FileClose($h)
EndFunc   ;==>_InitFiles_v1
Func _InitFiles_v2()
	Global $rootfolder = "openhabroot\"
	Local $h = FileOpen($rootfolder & "things\modbus_SH.things", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileWriteLine($h, "// Modbus to Smart-House OH3")
	FileClose($h)
	Global $hFileCfg = FileOpen($rootfolder & "things\modbus_SH.things", $FO_APPEND + $FO_UTF8_NOBOM)
	;
	Local $h = FileOpen($rootfolder & "items\smart-house.items", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileClose($h)
	Global $hFileItems = FileOpen($rootfolder & "items\smart-house.items", $FO_APPEND + $FO_UTF8_NOBOM)
	;
	Local $h = FileOpen($rootfolder & "rules\smart-house-sequences.rules-temp", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileClose($h)
	Global $hFileRules = FileOpen($rootfolder & "rules\smart-house-sequences.rules-temp", $FO_APPEND + $FO_UTF8_NOBOM)
	;
	Local $h = FileOpen($rootfolder & "items\smart-house-groups.items", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileWriteLine($h, "// Groups")
	FileClose($h)
	Global $hFileItemsGroups = FileOpen($rootfolder & "items\smart-house-groups.items", $FO_APPEND + $FO_UTF8_NOBOM)
	;
	Local $h = FileOpen($rootfolder & "transform\dimmertransform.js", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileWriteLine($h, "(function (inputData) {")
	FileWriteLine($h, "	   return inputData;")
	FileWriteLine($h, "})(input)")
	FileClose($h)
	;
	Local $h = FileOpen($rootfolder & "transform\divide10.js", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileWriteLine($h, "(function (inputData) {")
	FileWriteLine($h, "	   return parseFloat(inputData) / 10;")
	FileWriteLine($h, "})(input)")
	FileClose($h)
	;
	Local $h = FileOpen($rootfolder & "transform\multiply10.js", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
	FileWriteLine($h, "(function (inputData) {")
	FileWriteLine($h, "	   return Math.round(parseFloat(inputData, 10) * 10);")
	FileWriteLine($h, "})(input)")
	FileClose($h)
EndFunc   ;==>_InitFiles_v2

Func _AddGroup($x, $y)
	; Search file if group $x exists, if not add it
	Local $aGroupline
	If _FileReadToArray($rootfolder & "items\smart-house-groups.items", $aGroupline) Then
		Local $fExist = False
		For $i = 1 To UBound($aGroupline) - 1
			If StringInStr($aGroupline[$i], $x) Then
				; Group already exists
				$fExist = True
				ExitLoop
			EndIf
		Next
		If Not $fExist Then
			; add $x to file
			FileWriteLine($hFileItemsGroups, "Group|" & $x & "|" & Chr(34) & StringStripWS($y, 7) & Chr(34))
		EndIf
	EndIf
EndFunc   ;==>_AddGroup

Func _ExtractPlace($x)
	; Find places and functionnames
	; Place is between "(fx)" and "-"
	; Function is between "-" and "-" and in functionslist
	;
	; Remove "(Fx) "
	$x = StringReplace($x, "(fx) ", "")
	;
	; Look for Place
	$pos = StringInStr($x, " - ")
	If $pos > 0 Then
		Return StringLeft($x, $pos)
	Else
		Return ""
	EndIf
EndFunc   ;==>_ExtractPlace

Func _CloseFiles()
	FileClose($hFileCfg)
	FileClose($hFileItems)
	FileClose($hFileItemsGroups)
EndFunc   ;==>_CloseFiles

Func _FormatItemFile()
	; replace "|" with x spc's to have readable columns
	; ex Dimmer  HallDimmerBrunavggen  "(Fx) Hall - Dimmer Bruna väggen"  <slider>  (Dimbart ljus,Hall)  ["Lighting"]  {modbus="slave46:0"}
	;

	Local $nMaxItemtype = 0
	Local $nMaxName = 0
	Local $nMaxDescription = 0
	Local $nMaxIcon = 0
	Local $nMaxGroup = 0
	Local $nMaxHomekit = 0
	Local $nMaxChannel = 0
	Local $aFileItems
	Local $aFileGroups
	If _FileReadToArray($rootfolder & "items\smart-house-groups.items", $aFileGroups) Then
		_ArrayDelete($aFileGroups, 1)
		_FileWriteFromArray($rootfolder & "items\smart-house-groups.items", $aFileGroups, 1)
	Else
		ConsoleWrite("Fel på fil1:" & @error)
	EndIf
	If _FileReadToArray($rootfolder & "items\smart-house-groups.items", $aFileGroups, 0, "|") Then
		For $i = 1 To UBound($aFileGroups, 1) - 1
			$nMaxItemtype = _CountLen($aFileGroups[$i][0], $nMaxItemtype)
			$nMaxName = _CountLen($aFileGroups[$i][1], $nMaxName)
			$nMaxDescription = _CountLen($aFileGroups[$i][2], $nMaxDescription)
			If UBound($aFileGroups, 2) > 3 Then $nMaxIcon = _CountLen($aFileGroups[$i][3], $nMaxIcon)
			If UBound($aFileGroups, 2) > 4 Then $nMaxGroup = _CountLen($aFileGroups[$i][4], $nMaxGroup)
		Next
		FileClose($hFileItemsGroups)
		$hFileItemsGroups = FileOpen($rootfolder & "items\smart-house-groups.items", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
		For $i = 0 To UBound($aFileGroups, 1) - 1
			FileWrite($hFileItemsGroups, _AddSpc($aFileGroups[$i][0], $nMaxItemtype))
			FileWrite($hFileItemsGroups, _AddSpc($aFileGroups[$i][1], $nMaxName))
			FileWrite($hFileItemsGroups, _AddSpc($aFileGroups[$i][2], $nMaxDescription))
			If UBound($aFileGroups, 2) > 3 Then FileWrite($hFileItemsGroups, _AddSpc($aFileGroups[$i][3], $nMaxIcon))
			If UBound($aFileGroups, 2) > 3 Then FileWrite($hFileItemsGroups, _AddSpc($aFileGroups[$i][4], $nMaxGroup))
			FileWrite($hFileItemsGroups, @CRLF)
		Next
	Else
		ConsoleWrite("Fel på gruppfilen:" & @error)
	EndIf
	If _FileReadToArray($rootfolder & "items\smart-house.items", $aFileItems, 0, "|") Then
		;_ArrayDisplay($aFileItems)
		For $i = 0 To UBound($aFileItems, 1) - 1
			$nMaxItemtype = _CountLen($aFileItems[$i][0], $nMaxItemtype)
			$nMaxName = _CountLen($aFileItems[$i][1], $nMaxName)
			$nMaxDescription = _CountLen($aFileItems[$i][2], $nMaxDescription)
			$nMaxIcon = _CountLen($aFileItems[$i][3], $nMaxIcon)
			$nMaxGroup = _CountLen($aFileItems[$i][4], $nMaxGroup)
			$nMaxHomekit = _CountLen($aFileItems[$i][5], $nMaxHomekit)
			$nMaxChannel = _CountLen($aFileItems[$i][6], $nMaxChannel)
		Next
		FileClose($hFileItems)
		$hFileItems = FileOpen($rootfolder & "items\smart-house.items", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UTF8_NOBOM)
;~ 			FileWrite($hFileItems, @CRLF)
		For $i = 0 To UBound($aFileItems, 1) - 1
			FileWrite($hFileItems, _AddSpc($aFileItems[$i][0], $nMaxItemtype))
			FileWrite($hFileItems, _AddSpc($aFileItems[$i][1], $nMaxName))
			FileWrite($hFileItems, _AddSpc($aFileItems[$i][2], $nMaxDescription))
			FileWrite($hFileItems, _AddSpc($aFileItems[$i][3], $nMaxIcon))
			FileWrite($hFileItems, _AddSpc($aFileItems[$i][4], $nMaxGroup))
			FileWrite($hFileItems, _AddSpc($aFileItems[$i][5], $nMaxHomekit))
			FileWrite($hFileItems, _AddSpc($aFileItems[$i][6], $nMaxChannel))
			FileWrite($hFileItems, @CRLF)
		Next
;~ 		FileDelete($rootfolder & "items\smart-house-groups.items")
	Else
		MsgBox(Default, "Error", "Error: " & @error)
	EndIf
EndFunc   ;==>_FormatItemFile

Func _AddSpc($x, $max)
	Return $x & _StringRepeat(" ", $max - StringLen($x)) & " "
EndFunc   ;==>_AddSpc

Func _CountLen($x1, $x2)
	If StringLen($x1) > $x2 Then
		Return StringLen($x1)
	Else
		Return $x2
	EndIf
EndFunc   ;==>_CountLen

Func _ModbusCfgFile($amodbusfile)
;~ ####### Modbus to Smart-House ########

;~ # smart-house (Fx) Sovrum 4 (Kontor) - Ljusfunktion Tak_Status
;~ modbus:tcp.slave1.connection=192.168.1.198:502
;~ modbus:tcp.slave1.type=holding
;~ modbus:tcp.slave1.id=1
;~ modbus:tcp.slave1.start=0
;~ modbus:tcp.slave1.length=1
;~ modbus:tcp.slave1.valuetype=uint16
	Local $fFunctions = False ; Flag set to True when functions start
	Local $mslave = 0
	Local $tNameFunction = ""
	For $i = 0 To UBound($amodbusfile, 1) - 1
		;find first line for functions
		;Look row-wise for "IR" or "HR"
		If Not $fFunctions Then
			While 1 = 1
				For $col = 0 To UBound($amodbusfile, 2) - 1
					If StringInStr("IR HR", $amodbusfile[$i][$col]) Then ExitLoop 2
				Next
				$i += 1
			WEnd
			$fFunctions = True

			;find columns to use
			$col = -1
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colFxType = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colFxName = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModRegType = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModID = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModAdr = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModAdrH = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModType = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModScale = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colFxVar = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colFxUnit = $col
		EndIf
		$aFunction = _LookupItem($amodbusfile[$i][$colFxType])

		; Filter functions to create modbusslaves on
		If StringInStr("Ljusfunktion Dimbart ljus Temperaturzon", $aFunction[3]) Then
			$sNameOH = $aFunction[0]
			$sIconOH = $aFunction[1]
			$sNameHomekit = $aFunction[2]
			$sNameSH_SE = $aFunction[3]

			; Run only when FxName has changed fom previous row
			While $tNameFunction <> $amodbusfile[$i][$colFxName]
				If StringInStr("Ljusfunktion Dimbart ljus", $sNameSH_SE) Then Local $aSystemVariable[1] = ["Status"]
				If StringInStr("Temperaturzon", $sNameSH_SE) Then Local $aSystemVariable[2] = ["Temperatur rum", "Status"]
				For $ii = 0 To UBound($aSystemVariable, 1) - 1
					; find next System Variable
					While $amodbusfile[$i][$colFxVar] <> $aSystemVariable[$ii]
						$i += 1
						If $i = UBound($amodbusfile, 1) Then ExitLoop 3
					WEnd
					$mslave += 1
					$s = "modbus:tcp.slave" & $mslave & "."

					; Create part in modbus.cfg-file
					FileWriteLine($hFileCfg, "")
					FileWriteLine($hFileCfg, "# smart-house " & $amodbusfile[$i][$colFxName] & "_" & $aSystemVariable[$ii])
					FileWriteLine($hFileCfg, $s & "connection=" & $mIP)
					FileWriteLine($hFileCfg, $s & "type=" & _Registertype($amodbusfile[$i][$colModRegType]))
					FileWriteLine($hFileCfg, $s & "id=" & $amodbusfile[$i][$colModID])
					FileWriteLine($hFileCfg, $s & "start=" & $amodbusfile[$i][$colModAdr])
					FileWriteLine($hFileCfg, $s & "length=" & _Registerlength($amodbusfile[$i][$colModType]))
					FileWriteLine($hFileCfg, $s & "valuetype=" & _RegisterValueType($amodbusfile[$i][$colModType]))

					; Dimmer needs to be transformed, example:
					; {modbus=">[slave2:0],<[slave2:0:transformation=JS(identity.js)]"}

					; Temperature needs to be transformed, example:
					; {modbus="<[slave5:0:transformation=JS(divide10.js)]"}

					; Temperature needs to have variable "$FxNameVar" added to FXname

					;Change icon and Homekit type to switch/Switchable if it is a ReadWrite "Status" function ("läge" in name)
					If _itemstatusReadWrite($amodbusfile[$i][$colFxName]) Then
						$sIconOH = "switch"
						$sNameHomekit = "Switchable"
						$sNameSH_SE = "Läge"
					EndIf

					;Change icon to switch and HomeKit to "" if it is a Readonly "Status" function ("status" in name)
					If _itemstatusRead($amodbusfile[$i][$colFxName]) Then
						$sIconOH = "switch"
						$sNameHomekit = ""
						$sNameSH_SE = "Status"
					EndIf


					$sFxNameVar = ""
					Select
						Case $sNameSH_SE = "Dimbart ljus"
							$sBinding = "{modbus=" & Chr(34) & ">[slave" & $mslave & ":0],<[slave" & $mslave & ":0:transformation=JS(dimmertransform.js)]" & Chr(34) & "}"
						Case $sNameSH_SE = "Temperaturzon" And $aSystemVariable[$ii] = "Temperatur rum"
							$sBinding = "{modbus=" & Chr(34) & "<[slave" & $mslave & ":0:transformation=JS(divide10.js)]" & Chr(34) & "}"
							$sFxNameVar = " [%.1f °C]"
						Case $sNameSH_SE = "Temperaturzon" And $aSystemVariable[$ii] = "Status"
							$sBinding = "{modbus=" & Chr(34) & "slave" & $mslave & ":0" & Chr(34) & "}"
							$sNameOH = "Contact"
							$sIconOH = "switch"
							$sNameHomekit = ""
						Case Else
							$sBinding = "{modbus=" & Chr(34) & "slave" & $mslave & ":0" & Chr(34) & "}"
					EndSelect
					;
					; Create Line in group-file
					$sNamePlace = ""
					$sPlace = _ExtractPlace($amodbusfile[$i][1])
					If $sPlace <> "" Then
						$sNamePlace = "," & _itemname($sPlace)

						_AddGroup(_itemname($sPlace), $sPlace)
					EndIf
					_AddGroup(_itemname($sNameSH_SE), $sNameSH_SE)

					; Create line in Itemsfile
					; Ex Dimmer|HallDimmerBrunavggen|"(Fx) Hall - Dimmer Bruna väggen"|<slider>|(Dimbart ljus,Hall)|["Lighting"]|{modbus="slave46:0"}
					FileWriteLine($hFileItems, _
							$sNameOH & "|" _
							 & _itemname($amodbusfile[$i][$colFxName] & "_" & $aSystemVariable[$ii]) & "|" _
							 & Chr(34) & StringReplace($amodbusfile[$i][$colFxName], "(Fx) ", "") & $sFxNameVar & Chr(34) & "|<" _
							 & $sIconOH & ">|" _
							 & "(" & _itemname($sNameSH_SE) & $sNamePlace & ")|" _
							 & "[" & Chr(34) & $sNameHomekit & Chr(34) & "]|" _
							 & $sBinding)

				Next
				;				$tNameFunction = $amodbusfile[$i][1]
				$tNameFunction = $amodbusfile[$i][$colFxName]
				$i += 1

			WEnd
		EndIf
	Next
EndFunc   ;==>_ModbusCfgFile
Func _ModbusThingsFile($amodbusfile)
;~ // Modbus to Smart-House

;~ Bridge modbus:tcp:SH2WEB [ host="192.168.1.198", port=502, id=1 ] {

;~ // smart-house (Fx) Sovrum 4 (Kontor) - Ljusfunktion Tak_Status
;~ 	Bridge poller "name1" [ start=0, length=4, refresh=1000, type="holding" ] {
;~         Thing data "name1" [ readStart="0", readValueType="uint16", writeStart="0", writeValueType="uint16", writeType="holding" ]
;~ 	Bridge poller "name2" [ start=2, length=4, refresh=1000, type="holding" ] {
;~         Thing data "name2" [ readStart="2", readValueType="uint16", writeStart="2", writeValueType="uint16", writeType="holding" ]
;~     }
;~ }

	;	Parse the csv-file
	FileWriteLine($hFileCfg, "Bridge modbus:tcp:SH2WEB [ host=" & Chr(34) & $mIP & Chr(34) & ", port=" & $mPort & ", id=" & $mID & " ] {")
	Local $fFunctions = False ; Flag set to True when functions start
	Local $mslave = 0
	Local $tNameFunction = ""
	For $i = 0 To UBound($amodbusfile, 1) - 1
		;find first line for functions
		;Look row-wise for "IR" or "HR"
		If Not $fFunctions Then
			While 1 = 1
				For $col = 0 To UBound($amodbusfile, 2) - 1
					If StringInStr("IR HR", $amodbusfile[$i][$col]) Then ExitLoop 2
				Next
				$i += 1
			WEnd
			$fFunctions = True

			;find columns to use
			$col = -1
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colFxType = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colFxName = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModRegType = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModID = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModAdr = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModAdrH = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModType = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colModScale = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colFxVar = $col
			Do
				$col += 1
			Until $amodbusfile[$i][$col] <> ""
			$colFxUnit = $col
		EndIf


		; Run only when FxName has changed fom previous row
		While $tNameFunction <> $amodbusfile[$i][$colFxName]
		$aFunction = _LookupItem($amodbusfile[$i][$colFxType])
			; Filter functions to create pollers on
			If StringInStr("Ljusfunktion Dimbart ljus Switch Function Temperaturzon Analog komparator", $aFunction[3]) Then
				$sNameOH = $aFunction[0]
				$sIconOH = $aFunction[1]
				$sNameHomekit = $aFunction[2]
				$sNameSH_SE = $aFunction[3]
				If StringInStr("Ljusfunktion Dimbart ljus Switch Function Analog komparator", $sNameSH_SE) Then Local $aSystemVariable[1] = ["Status"]
				If StringInStr("Temperaturzon", $sNameSH_SE) Then Local $aSystemVariable[2] = ["Temperatur rum", "Status"]
				For $ii = 0 To UBound($aSystemVariable, 1) - 1
					; find next System Variable
					While $amodbusfile[$i][$colFxVar] <> $aSystemVariable[$ii]

						$i += 1
						If $i = UBound($amodbusfile, 1) Then ExitLoop 3
					WEnd
					$mslave += 1

					;**********************************************************************************************************************************************
					; Flyttade följande rader. Nu blir det fel i items-filen m.m. Analog komparator verkar komma med som ljusfunktion??, se rad 94 i smart-house.items
					;**********************************************************************************************************************************************

					;Change icon and Homekit type to switch/Switchable if it is a ReadWrite "Status" function ("läge" in name)
					If _itemstatusReadWrite($amodbusfile[$i][$colFxName]) Then
						$sIconOH = "switch"
						$sNameHomekit = "Switchable"
						$sNameSH_SE = "Läge"
					EndIf

					;Change icon to switch and HomeKit to "" if it is a Readonly "Status" function ("status" in name)
					If _itemstatusRead($amodbusfile[$i][$colFxName]) Then
						$sIconOH = "switch"
						$sNameHomekit = ""
						$sNameSH_SE = "Status"
					EndIf

					If _Registertype($amodbusfile[$i][$colModRegType]) = "holding" And $sNameSH_SE <> "Status" And Not StringInStr("Temperaturzon", $sNameSH_SE) Then
						$fReadOnly = False
						$fRefresh = 1000
					Else
						$fReadOnly = True
						$fRefresh = 10000
					EndIf


					; Create part in .things-file
					FileWriteLine($hFileCfg, "    // smart-house " & $amodbusfile[$i][$colFxName] & "_" & $aSystemVariable[$ii])

					$sPoller = "    Bridge poller P" & $mslave
					$sPoller = $sPoller & " [ start=" & $amodbusfile[$i][$colModAdr]
					$sPoller = $sPoller & ", length=" & _Registerlength($amodbusfile[$i][$colModType])
					$sPoller = $sPoller & ", refresh=" & $fRefresh
					$sPoller = $sPoller & ", type=" & Chr(34) & _Registertype($amodbusfile[$i][$colModRegType]) & Chr(34) & " ] {"
					FileWriteLine($hFileCfg, $sPoller)

					$sData = "        Thing data D" & $mslave
					$sData = $sData & " [ readStart=" & Chr(34) & $amodbusfile[$i][$colModAdr] & Chr(34)
					$sData = $sData & ", readValueType=" & Chr(34) & _RegisterValueType($amodbusfile[$i][$colModType]) & Chr(34)

					;readTransform="JS(divide10.js)", 1 decimal on temperature
					If $sNameSH_SE = "Temperaturzon" And $aSystemVariable[$ii] = "Temperatur rum" Then
						$sData = $sData & ", readTransform=" & Chr(34) & "JS(divide10.js)" & Chr(34)
					EndIf

					;readTransform="JS(dimmertransform.js)", Transform percentage of dimmer to KEEP as percentage, discussed here: https://community.openhab.org/t/modbus-openhab2-binding-available-for-alpha-testing/27657/551?u=pe_man
					If $sNameSH_SE = "Dimbart ljus" Then
						$sData = $sData & ", readTransform=" & Chr(34) & "JS(dimmertransform.js)" & Chr(34)
					EndIf
					If Not $fReadOnly Then
						$sData = $sData & ", writeStart=" & Chr(34) & $amodbusfile[$i][$colModAdr] & Chr(34)
						$sData = $sData & ", writeValueType=" & Chr(34) & _RegisterValueType($amodbusfile[$i][$colModType]) & Chr(34)
						$sData = $sData & ", writeType=" & Chr(34) & _Registertype($amodbusfile[$i][$colModRegType]) & Chr(34)
					EndIf
					$sData = $sData & " ]"
					FileWriteLine($hFileCfg, $sData)
					FileWriteLine($hFileCfg, "    }")

					; Dimmer needs to be transformed, example:
					; {modbus=">[slave2:0],<[slave2:0:transformation=JS(identity.js)]"}

					; Temperature needs to be transformed, example:
					; {modbus="<[slave5:0:transformation=JS(divide10.js)]"}

					; Temperature needs to have variable "$FxNameVar" added to FXname


					$sFxNameVar = ""
					Select
						Case $sNameSH_SE = "Dimbart ljus"
							$sBinding = "{ channel=" & Chr(34) & "modbus:data:SH2WEB:P" & $mslave & ":D" & $mslave & ":dimmer" & Chr(34) & " }"
						Case $sNameSH_SE = "Temperaturzon" And $aSystemVariable[$ii] = "Temperatur rum"
							$sBinding = "{ channel=" & Chr(34) & "modbus:data:SH2WEB:P" & $mslave & ":D" & $mslave & ":number" & Chr(34) & " }"
							$sFxNameVar = " [%.1f °C]"
						Case $sNameSH_SE = "Temperaturzon" And $aSystemVariable[$ii] = "Status"
							$sBinding = "{ channel=" & Chr(34) & "modbus:data:SH2WEB:P" & $mslave & ":D" & $mslave & ":contact" & Chr(34) & " }"
							$sNameOH = "Contact"
							$sIconOH = "switch"
							$sNameHomekit = ""
						Case $sNameSH_SE = "Status"
							$sBinding = "{ channel=" & Chr(34) & "modbus:data:SH2WEB:P" & $mslave & ":D" & $mslave & ":contact" & Chr(34) & " }"
							$sNameOH = "Contact"
							$sIconOH = "switch"
							$sNameHomekit = ""

						Case Else
							$sBinding = "{ channel=" & Chr(34) & "modbus:data:SH2WEB:P" & $mslave & ":D" & $mslave & ":switch" & Chr(34) & " }"
					EndSelect
					;
					; Create Line in group-file
					$sNamePlace = ""
					$sPlace = _ExtractPlace($amodbusfile[$i][1])
					If $sPlace <> "" Then
						$sNamePlace = "," & _itemname($sPlace)

						_AddGroup(_itemname($sPlace), $sPlace)
					EndIf
					_AddGroup(_itemname($sNameSH_SE), $sNameSH_SE)

					; Create line in Itemsfile
					; Ex Dimmer|HallDimmerBrunavggen|"(Fx) Hall - Dimmer Bruna väggen"|<slider>|(Dimbart ljus,Hall)|["Lighting"]|{modbus="slave46:0"}
					FileWriteLine($hFileItems, _
							$sNameOH & "|" _
							 & _itemname($amodbusfile[$i][$colFxName] & "_" & $aSystemVariable[$ii]) & "|" _
							 & Chr(34) & StringReplace($amodbusfile[$i][$colFxName], "(Fx) ", "") & $sFxNameVar & Chr(34) & "|<" _
							 & $sIconOH & ">|" _
							 & "(" & _itemname($sNameSH_SE) & $sNamePlace & ")|" _
							 & "[" & Chr(34) & $sNameHomekit & Chr(34) & "]|" _
							 & $sBinding)

				Next
			EndIf
				$tNameFunction = $amodbusfile[$i][$colFxName]
				$i += 1

		WEnd
	Next
	FileWriteLine($hFileCfg, "}")
EndFunc   ;==>_ModbusThingsFile

Func _Registerlength($x)
	Return (StringRight($x, 2) / 16)
EndFunc   ;==>_Registerlength

Func _RegisterValueType($x)
	If _Registerlength($x) = 2 Then
		Return StringLower($x) & "_swap"
	Else
		Return StringLower($x)
	EndIf
EndFunc   ;==>_RegisterValueType

Func _itemname($x)
	; Wash Functionname from illegal and unwanted characters
	If StringLeft($x, 5) = "(Fx) " Then $x = StringRight($x, StringLen($x) - 5)
	$x = StringReplace($x, "-", " ")
	$x = StringStripWS($x, 7)
	Return StringRegExpReplace($x, "[[:^alnum:]]", "_")
EndFunc   ;==>_itemname
Func _itemdisplay($x)
EndFunc   ;==>_itemdisplay
Func _itemplace($x)
	; identify a Place. First word(s) separated with "-"
	; Returns Place or ""
	$ax = StringSplit($x, "-")
	If UBound($ax) > 0 Then
		Return _itemname($ax[1])
	Else
		Return ""
	EndIf
EndFunc   ;==>_itemplace
Func _itemstatusRead($x)
	; Identify if the function is a Read-Only status-function from its name
	; Return true if Status is found in name
	$result = False
	$pos = StringInStr($x, "status")
	; $xmod = $x                                      TA BORT!
	If $pos > 0 Then $result = True
	Return $result
EndFunc   ;==>_itemstatusRead
Func _itemstatusReadWrite($x)
	; Identify if the function is a writable status-function from its name
	; Return true if läge is found in name
	$result = False
	$pos = StringInStr($x, "läge")
	; $xmod = $x                                      TA BORT!
	If $pos > 0 Then $result = True
	Return $result
EndFunc   ;==>_itemstatusReadWrite
Func _Registertype($x)
	If $x = "IR" Then Return "input"
	If $x = "HR" Then Return "holding"
EndFunc   ;==>_Registertype
Func _LookupItem($x)
	; Find wich item to convert to
	; Itemtype, icon, Homekittype, Smarthousefunctiontype (swe), Smarthousefunctiontype (eng)
	Local $aResult[] = ["", "", "", "", ""]
	Local $aArray[][] = [ _
			["Switch", "light", "Lighting", "Ljusfunktion", "Light function"], _
			["Switch", "switch", "Switchable", "Switch Function", "Switch Function"], _
			["Dimmer", "slider", "Lighting", "Dimbart ljus", "Dimmable light"], _
			["Number", "", "", "Räknarfunktion", "Counter function"], _
			["Number", "time", "", "Fördröjningstimer", "Delay timer"], _
			["Number", "time", "", "Intervalltimer", "Interval timer"], _
			["Number", "", "", "Inbrott huvudlarm", "Main intruder alarm"], _
			["Number", "", "", "Matematisk funktion", "Mathematical function"], _
			["Contact", "switch", "", "Multigate", "Multigate"], _
			["Number", "time", "", "Pulserande timer", "Recycling timer"], _
			["Switch", "switch", "Switchable", "Sekvens", "Sequence"], _
			["", "", "", "Simulerat boende", "Simulated habitation"], _
			["", "", "", "Siren", "Siren"], _
			["", "", "", "Brandvarnare", "Smoke alarm"], _
			["", "", "", "Systemfunktion", "Systemfunktion"], _
			["", "", "", "Inbrott larmzon", "Zone intruder alarm"], _
			["Number", "temperature", "CurrentTemperature", "Temperaturzon", "Zone temperature"], _
			["Contact", "switch", "", "Analog komparator", "Analog comparator"], _
			["", "", "", "Motorvärmare", "Car heating"], _
			["Number", "time", "", "Timräknare", "Hour counter"]]
	;_ArrayDisplay($aArray)
	$result = _ArraySearch($aArray, $x, 0, 0, 0, 0, 1, 3)
	If $result <> -1 Then
		For $i = 0 To 4
			$aResult[$i] = $aArray[$result][$i]
		Next
	EndIf
	Return $aResult
EndFunc   ;==>_LookupItem
Func VisaFil($f)
	; Define a variable to pass to _FileReadToArray.
	Local $aArray = 0

	; Read the current script file into an array using the variable defined previously.
	; $iFlag is specified as 0 in which the array count will not be defined. Use UBound() to find the size of the array.
	If Not _FileReadToArray($f, $aArray, 0, ";") Then
		MsgBox($MB_SYSTEMMODAL, "", "There was an error reading the file. @error: " & @error) ; An error occurred reading the current script file.
	EndIf

	; Display the array in _ArrayDisplay.
	;_ArrayDisplay($aArray)
	Return $aArray
EndFunc   ;==>VisaFil
Func Choose_file()
	; Create a constant variable in Local scope of the message to display in FileOpenDialog.
	Local Const $sMessage = "Hold down Ctrl or Shift to choose multiple files."

	; Display an open dialog to select a list of file(s).
	Local $sFileOpenDialog = FileOpenDialog($sMessage, @WindowsDir & "\", "CSV (*.csv)", $FD_FILEMUSTEXIST + $FD_MULTISELECT)
	If @error Then
		; Display the error message.
		MsgBox($MB_SYSTEMMODAL, "", "No file(s) were selected.")

		; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
		FileChangeDir(@ScriptDir)
	Else
		; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
		FileChangeDir(@ScriptDir)
		Return $sFileOpenDialog
		; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
		;$sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)

		; Display the list of selected files.
		;MsgBox($MB_SYSTEMMODAL, "", "You chose the following files:" & @CRLF & $sFileOpenDialog)
	EndIf
EndFunc   ;==>Choose_file
