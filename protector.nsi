  !include "MUI2.nsh"
  !include "FileAssociation.nsh"

  Name "Windows Protec'tor 1.0"
  !define VERSION "1.0"
  !define NAME "protector"
  !define UNINSTREG "Software\Microsoft\Windows\CurrentVersion\Uninstall\ctprotector"
  !define ENABLE_LOGGING
  OutFile "ct_protector.exe"
  RequestExecutionLevel admin


!define LogSet "!insertmacro LogSetMacro"
!macro LogSetMacro SETTING
  !ifdef ENABLE_LOGGING
    LogSet ${SETTING}
  !endif
!macroend

!define LogText "!insertmacro LogTextMacro"
!macro LogTextMacro INPUT_TEXT
  !ifdef ENABLE_LOGGING
    LogText ${INPUT_TEXT}
  !endif
!macroend

!macro KeyExists ROOT MAIN_KEY KEY
  Push $R0
  Push $R1
  Push $R2


  StrCpy $R1 "0" # loop index
  StrCpy $R2 "0" # not found

  ${Do}
    EnumRegKey $R0 ${ROOT} "${MAIN_KEY}" "$R1"
    ${If} $R0 == "${KEY}"
      StrCpy $R2 "1" # found
      ${Break}
    ${EndIf}
    IntOp $R1 $R1 + 1
  ${LoopWhile} $R0 != ""

  ClearErrors

  Exch 2
  Pop $R0
  Pop $R1
  Exch $R2
!macroend



!define LVM_GETITEMTEXT 0x102D

Function DumpLog
  Exch $5
  Push $0
  Push $1
  Push $2
  Push $3
  Push $4
  Push $6

  FindWindow $0 "#32770" "" $HWNDPARENT
  GetDlgItem $0 $0 1016
  StrCmp $0 0 exit
  FileOpen $5 $5 "w"
  StrCmp $5 "" exit
    SendMessage $0 ${LVM_GETITEMCOUNT} 0 0 $6
    System::Alloc ${NSIS_MAX_STRLEN}
    Pop $3
    StrCpy $2 0
    System::Call "*(i, i, i, i, i, i, i, i, i) i \
      (0, 0, 0, 0, 0, r3, ${NSIS_MAX_STRLEN}) .r1"
    loop: StrCmp $2 $6 done
      System::Call "User32::SendMessageA(i, i, i, i) i \
        ($0, ${LVM_GETITEMTEXT}, $2, r1)"
      System::Call "*$3(&t${NSIS_MAX_STRLEN} .r4)"
      FileWrite $5 "$4$\r$\n"
      IntOp $2 $2 + 1
      Goto loop
    done:
      FileClose $5
      System::Free $1
      System::Free $3
  exit:
    Pop $6
    Pop $4
    Pop $3
    Pop $2
    Pop $1
    Pop $0
    Exch $5
FunctionEnd


!macro IsAdmin
UserInfo::GetAccountType
pop $0
${If} $0 != "admin" ;Administratorrechte einfordern
        messageBox mb_iconstop "Windows Protec'tor muss mit Administratorrechten ausgeführt werden."
        setErrorLevel 740 ;Keine Ausführung ohne Administratorrechte
        quit
${EndIf}
!macroend

         !define MUI_ABORTWARNING "Hiermit beenden Sie die Einrichtung von c't Windows Protec'tor"
         !define MUI_ICON "ct.ico"
         !define MUI_WELCOMEPAGE_TITLE "Windows Protec'tor"
         !define MUI_WELCOMEPAGE_TEXT "Mit Windows Protec'tor sichern Sie Windows, Office und den Acrobat Reader ab. Dazu werden vor allem Funktionen deaktiviert, auf die viele Benutzer verzichten können. Auf der nächsten Seite haben Sie die Möglichkeit, die für Sie geeigneten Maßnahmen auszuwählen. $\r$\n$\r$\nInformationen und Updates unter https://www.ct.de/protector"
         !define MUI_UNTEXT_FINISH_TITLE "Abgeschlossen."
         !define MUI_UNTEXT_FINISH_SUBTITLE ""
         !define MUI_UNTEXT_ABORT_TITLE "Abgebrochen"
         !define MUI_UNTEXT_ABORT_SUBTITLE ""

         !define MUI_PAGE_HEADER_TEXT "Maßnahmen auswählen"
         !define MUI_PAGE_HEADER_SUBTEXT "Funktionen deaktivieren, um die Sicherheit zu erhöhen."
         !define MUI_COMPONENTSPAGE_TEXT_TOP "Lesen Sie die Beschreibungen gründlich. Ein Haken sorgt dafür, dass eine Härtungsmaßnahme aktiviert wird."
         !define MUI_COMPONENTSPAGE_TEXT_COMPLIST "Alle Härtungs-Maßnahmen"

        !define MUI_COMPONENTSPAGE_TEXT_INSTTYPE ""

        !define MUI_COMPONENTSPAGE_TEXT_DESCRIPTION_TITLE "Beschreibung der Maßnahme"
        !define MUI_COMPONENTSPAGE_TEXT_DESCRIPTION_INFO "Fahren Sie mit der Maus über die Maßnahmen, um Details zu sehen."

InstallDir "$PROGRAMFILES\ctProtector"

!insertmacro MUI_LANGUAGE "German"


  VIAddVersionKey /LANG=${LANG_GERMAN} "ProductName" "Windows Protec'tor"
VIAddVersionKey /LANG=${LANG_GERMAN} "Comments" ""
VIAddVersionKey /LANG=${LANG_GERMAN} "CompanyName" "c't Magazin für Computertechnik"
VIAddVersionKey /LANG=${LANG_GERMAN} "LegalTrademarks" ""
VIAddVersionKey /LANG=${LANG_GERMAN} "LegalCopyright" "GNU General Public License v3.0"
VIAddVersionKey /LANG=${LANG_GERMAN} "FileDescription" ""
VIAddVersionKey /LANG=${LANG_GERMAN} "FileVersion" "1.0"
VIProductVersion 1.0.0.0


var beforeScriptHost
var beforeAutorun1
var beforeAutorun2
var beforeAutorun3
var beforeUAC
var beforeShowFile

var officeVer

var beforeMakroWord
var beforeMakroExcel
var beforeMakroPowerpoint
var beforeOfficeActiveX

var beforeOLEWord
var beforeOLEExcel
var beforeOLEPowerpoint

var beforeReaderXIX
var beforeReaderDCX
var beforeReaderXIObject
var beforeReaderDCObject

var isset




  !insertmacro MUI_PAGE_WELCOME
  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_INSTFILES
  
  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES
  
  var sh
var ar1
var ar2
var ar3
var uac
var fex

!macro setDefaults un

Function ${un}setDefaults

ReadRegDWORD $sh HKCU "SOFTWARE\CT Protector\" "Scripthost"
        ${If} ${Errors}
        ${Else}
                WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows Script Host\Settings" "Enabled" $sh
        ${EndIf}

        ClearErrors

        ReadRegDWORD $ar1 HKCU "SOFTWARE\CT Protector\" "Autorun1"

        ${If} ${Errors}

        ${Else}
               WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoDriveTypeAutoRun" $ar1
        ${EndIf}

        ClearErrors

        ReadRegDWORD $ar2 HKCU "SOFTWARE\CT Protector\" "Autorun2"

         ${If} ${Errors}

        ${Else}
               WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoAutorun" $ar2
          ${EndIf}
              ClearErrors

            ReadRegDWORD $ar3 HKCU "SOFTWARE\CT Protector\" "Autorun3"

            ${If} ${Errors}

            ${Else}

            WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" "DisableAutoplay" $ar3

            ${EndIf}



        DeleteRegValue HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" "55"
        DeleteRegValue HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" "56"
        DeleteRegValue HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" "57"

        ClearErrors
        ReadRegDWORD $uac  HKCU "SOFTWARE\CT Protector\" "UAC"

                ${If} ${Errors}

                ${Else}

                       WriteRegDWORD HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" "ConsentPromptBehaviorAdmin" $uac
                ${EndIf}

        ${registerExtension} "" ".hta" "htafile"
        ${registerExtension} "" ".js" "JSFile"
        ${registerExtension} "" ".JSE" "JSEFile"
        ${registerExtension} "" ".WSH" "WSHFile"
        ${registerExtension} "" ".WSF" "WSFFile"
        ${registerExtension} "" ".scr" "scrfile"
        ${registerExtension} "" ".vbs" "VBSFile"
        ${registerExtension} "" ".pif" "piffile"
        ClearErrors

         ReadRegDWORD $fex HKCU "SOFTWARE\CT Protector\" "ShowFile"


                ${If} ${Errors}
                ${Else}
                      WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt" $fex

                ${EndIf}
FunctionEnd

!macroend
  
  
  !insertmacro setDefaults ""
  
  

; Dieser Abschnitt wird nicht zur Auswahl angezeigt:
SectionGroup "-sonstiges"
            Section "-Uninstaller" uninstall
             SectionIn RO
             setOutPath $INSTDIR
             writeUninstaller "uninstall.exe"


             CreateDirectory "$SMPROGRAMS\c't Windows Protec'tor"
             CreateShortCut "$SMPROGRAMS\c't Windows Protec'tor\Deinstallieren.lnk" "$INSTDIR\uninstall.exe"
             
            ; Damit die Deinstallation in der Systemsteuerung auftaucht
            WriteRegStr HKLM "${UNINSTREG}" "DisplayName" "${NAME}"
            WriteRegStr HKLM "${UNINSTREG}" "DisplayVersion" "${VERSION}"
            WriteRegStr HKLM "${UNINSTREG}" "Publisher" "c't-Redaktion"
            WriteRegStr HKLM "${UNINSTREG}" "URLInfoAbout" "http://www.ct.de/yt8z"
            WriteRegStr HKLM "${UNINSTREG}" "UninstallString" '"$INSTDIR\uninstall.exe"'
            WriteRegDWORD HKLM "${UNINSTREG}" "NoModify" 1
            WriteRegDWORD HKLM "${UNINSTREG}" "NoRepair" 1

             SectionEnd
             
        Section "-Cleanup" cleanup
             
             Call setDefaults
             
             SectionEnd
             
SectionGroupEnd




SectionGroup "Windows"

        Section "Windows-Script-Host abschalten" ScriptHost
        
               ;${LogText} "SomeTextHere"
        
                WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "Scripthost" $beforeScriptHost
                ${IF} $beforeScriptHost == 1
                      DetailPrint "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings\Enablded auf 0 gesetzt"
                      WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows Script Host\Settings" "Enabled" 0
                       DetailPrint "Windows Script Host wurde deaktiviert."
                ${ELSE}
                       DetailPrint "Windows Script Host ist bereits deaktiviert."
                ${ENDIF}
        SectionEnd

        Section "AutoRun und AutoPlay abschalten" AutoPlay

                         WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "Autorun1" $beforeAutorun1
                        ${IF} $beforeAutorun1 != 181
                              
                              DetailPrint "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDriveTypeAutoRun auf 181 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoDriveTypeAutoRun" 181
                               DetailPrint "AutoRun für alle Laufwerke deaktiviert."
                        ${ELSE}
                               DetailPrint "AutoRun ist bereits für alle Laufwerke deaktiviert."
                        ${ENDIF}
                
                        WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "Autorun2" $beforeAutorun2
                        ${IF} $beforeAutorun2 != 1
                              
                              DetailPrint "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoAutorun auf 1 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoAutorun" 1
                               DetailPrint "AutoRun deaktiviert."
                        ${ELSE}
                               DetailPrint "AutoRun ist bereits deaktiviert."
                        ${ENDIF}

                        WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "Autorun3" $beforeAutorun3
                        ${IF} $beforeAutorun3 != 1
                              
                              DetailPrint "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\DisableAutoplay auf 1 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" "DisableAutoplay" 1
                               DetailPrint "AutoPlay deaktiviert."
                        ${ELSE}
                               DetailPrint "AutoPlay ist bereits deaktiviert."
                        ${ENDIF}
     

        SectionEnd
        
        Section "PowerShell und CMD verhindern" PowerShell

                WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "PowerShell" "1"
                DetailPrint "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun cmd.exe hinzugefügt"
                DetailPrint "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun powershell.exe hinzugefügt"
                DetailPrint "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun powershell_ise.exe hinzugefügt"
        
                WriteRegStr HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" "55" "cmd.exe"
                WriteRegStr HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" "56" "powershell.exe"
                WriteRegStr HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" "57" "powershell_ise.exe"
                DetailPrint "cmd.exe und PowerShell sind deaktiviert."


        SectionEnd

        Section "User Account Control aktivieren" UAC
        
                              WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "UAC" $beforeUAC
                        ${IF} $beforeUAC != 181
                              
                              DetailPrint "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\ConsentPromptBehaviorAdmin auf 2 gesetzt"
                              WriteRegDWORD HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" "ConsentPromptBehaviorAdmin" 2
                               DetailPrint "UAC auf höchster Stufe aktiviert."
                        ${ELSE}
                               DetailPrint "UAC war bereits auf höchster Stufe."
                        ${ENDIF}

        SectionEnd


        Section "Bestimmte Dateiendungen nicht ausführen" FileExtensions
        
         WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "FileExtensions" "1"

        ${registerExtension} "" ".hta" ""
        ${registerExtension} "" ".js" ""
        ${registerExtension} "" ".JSE" ""
        ${registerExtension} "" ".WSH" ""
        ${registerExtension} "" ".WSF" ""
        ${registerExtension} "" ".scr" ""
        ${registerExtension} "" ".vbs" ""
        ${registerExtension} "" ".pif" ""
        
         ClearErrors
          DetailPrint "js, scr und weitere Dateiendungen deaktiviert."
        SectionEnd
        
              Section "Alle Dateiendungen anzeigen" FileShowExtensions

                             WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "ShowFile" $beforeShowFile
                        ${IF} $beforeShowFile == 1
                                  
                             DetailPrint "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt auf 0 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt" 0
                              DetailPrint "Dateiendungen werden immer angezeigt."
                              
                        ${ELSE}
                               DetailPrint "Dateiendungen wurden bereits angezeigt."
                        ${ENDIF}

        SectionEnd

SectionGroupEnd


SectionGroup "Microsoft Office" office

        Section "Word-Makros abschalten" WordMakros

                             WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "WordMakro" $beforeMakroWord
                             
                        ${IF} $beforeMakroWord != 4

                              
                              DetailPrint "HKCU\SOFTWARE\Microsoft\Office\$officeVer\Word\Security\VBAWarnings auf 4 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Office\$officeVer\Word\Security" "VBAWarnings" 4
                              DetailPrint "Makros deaktiviert in Word."

                        ${ELSE}
                               DetailPrint "Makros waren bereits deaktiviert in Word."
                        ${ENDIF}
        ClearErrors

         SectionEnd
         Section "Excel-Makros abschalten" ExcelMakros
         

                                 WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "ExcelMakro" $beforeMakroExcel
                        ${IF} $beforeMakroExcel != 4

                              DetailPrint "HKCU\SOFTWARE\Microsoft\Office\$officeVer\Excel\Security\VBAWarnings auf 4 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Office\$officeVer\Excel\Security" "VBAWarnings" 4
                              DetailPrint "Makros deaktiviert in Excel."

                        ${ELSE}
                               DetailPrint "Makros waren bereits deaktiviert in Excel."
                        ${ENDIF}

        ClearErrors

         SectionEnd
         Section "PowerPoint-Makros abschalten" PowerPointMakros
        
                                     WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "PowerPointMakro" $beforeMakroPowerPoint
                        ${IF} $beforeMakroPowerPoint != 4

                             
                              DetailPrint "HKCU\SOFTWARE\Microsoft\Office\$officeVer\PowerPoint\Security\VBAWarnings auf 4 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Office\$officeVer\PowerPoint\Security" "VBAWarnings" 4
                              DetailPrint "Makros deaktiviert in PowerPoint."

                        ${ELSE}
                               DetailPrint "Makros waren bereits deaktiviert in PowerPoint Version."
                        ${ENDIF}
        ClearErrors

        SectionEnd

        Section "OLE abschalten" OLE

                        WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "WordOLE" $beforeOLEWord
                        
                        ${IF} $beforeOLEWord != 2
                              DetailPrint "HKCU\SOFTWARE\Microsoft\Office\$officeVer\Word\Security\PackagerPrompt auf 2 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Office\$officeVer\Word\Security" "PackagerPrompt" 2
                              DetailPrint "OLEs deaktiviert in Word Version."

                        ${ELSE}
                               DetailPrint "OLEs waren bereits deaktiviert in Word Version."
                        ${ENDIF}

        ClearErrors

                          WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "ExcelOLE" $beforeOLEExcel
                        ${IF} $beforeOLEExcel != 2
                              DetailPrint "HKCU\SOFTWARE\Microsoft\Office\$officeVer\Excel\Security\PackagerPrompt auf 2 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Office\$officeVer\Excel\Security" "PackagerPrompt" 2
                              DetailPrint "OLEs deaktiviert in Excel Version."

                        ${ELSE}
                               DetailPrint "OLEs waren bereits deaktiviert in Excel Version."
                        ${ENDIF}

        ClearErrors

                         WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "PowerPointOLE" $beforeOLEPowerPoint
                        ${IF} $beforeOLEPowerPoint != 2
                              DetailPrint "HKCU\SOFTWARE\Microsoft\Office\$officeVer\PowerPoint\Security\PackagerPrompt auf 2 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Office\$officeVer\PowerPoint\Security" "PackagerPrompt" 2
                              DetailPrint "OLEs deaktiviert in PowerPoint Version."

                        ${ELSE}
                               DetailPrint "OLEs waren bereits deaktiviert in PowerPoint Version."
                        ${ENDIF}
        ClearErrors

        SectionEnd

        Section "ActiveX abschalten" ActiveX

                         WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "OfficeX" $beforeOfficeActiveX
                        ${IF} $beforeOfficeActiveX != 1
                              DetailPrint "HKCU\SOFTWARE\Microsoft\Office\Common\Security\DisableAllActiveX auf 1 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Microsoft\Office\Common\Security" "DisableAllActiveX" 1
                              DetailPrint "ActiveX in Office deaktiv."

                        ${ELSE}
                               DetailPrint "ActiveX in Office war bereits deaktiv."
                        ${ENDIF}
        ClearErrors

        SectionEnd

SectionGroupEnd

SectionGroup "Acrobat Reader"

        Section "JavaScript in PDF abschalten" ReaderJS

                            WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "ReaderDCX" $beforeReaderDCX
                        ${IF} $beforeReaderDCX == 1

                              DetailPrint "HKCU\SOFTWARE\Adobe\Acrobat Reader\DC\JSPrefs\bEnableJS auf 0 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Adobe\Acrobat Reader\DC\JSPrefs" "bEnableJS" 0
                              DetailPrint "ActiveX in Adobe Reader DC deaktiv."

                        ${ELSE}
                               DetailPrint "ActiveX war in Adobe Reader DC war bereits deaktiv."
                        ${ENDIF}
        ClearErrors
          
                           WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "ReaderXIX" $beforeReaderXIX
          
                        ${IF} $beforeReaderXIX == 1

                              DetailPrint "HKCU\SOFTWARE\Adobe\Acrobat Reader\XI\JSPrefs\bEnableJS auf 0 gesetzt"
                              WriteRegDWORD HKCU "SOFTWARE\Adobe\Acrobat Reader\XI\JSPrefs" "bEnableJS" 0
                              DetailPrint "ActiveX in Adobe Reader XI deaktiviert."

                        ${ELSE}
                               DetailPrint "ActiveX war in Adobe Reader XI war bereits deaktiv."
                        ${ENDIF}

        ClearErrors
          

        SectionEnd

        Section "Objektausführung in PDF verhindern" ReaderObjects

                 WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "ReaderXIObject" $beforeReaderXIObject

                        ${IF} $beforeReaderXIObject == 1
                              DetailPrint "HKCU\SOFTWARE\Adobe\Acrobat Reader\XI\Originals\bAllowOpenFile auf 0 gesetzt"
                              DetailPrint "HKCU\SOFTWARE\Adobe\Acrobat Reader\XI\Originals\bAllowOpenFile auf 0 gesetzt"
                              
                              WriteRegDWORD HKCU "SOFTWARE\Adobe\Acrobat Reader\XI\Originals" "bAllowOpenFile" 0
                              WriteRegDWORD HKCU "SOFTWARE\Adobe\Acrobat Reader\XI\Originals" "bSecureOpenFile" 1
                              DetailPrint "Objektausführung in Adobe Reader XI deaktiv."

                        ${ELSE}
                               DetailPrint "Objektausführung in Adobe Reader XI war bereits deaktiv."
                        ${ENDIF}
        ClearErrors
                        
                        WriteRegDWORD HKCU "SOFTWARE\CT Protector\" "ReaderDCObject" $beforeReaderDCObject
                        ${IF} $beforeReaderDCObject == 1
                              DetailPrint "HKCU\SOFTWARE\Adobe\Acrobat Reader\DC\Originals\bAllowFile auf 0 gesetzt"
                              DetailPrint "HKCU\SOFTWARE\Adobe\Acrobat Reader\DC\Originals\bSecureOpenFile auf 1 gesetzt"
                              
                              WriteRegDWORD HKCU "SOFTWARE\Adobe\Acrobat Reader\DC\Originals" "bAllowOpenFile" 0
                              WriteRegDWORD HKCU "SOFTWARE\Adobe\Acrobat Reader\DC\Originals" "bSecureOpenFile" 1
                              DetailPrint "Objektausführung in Adobe Reader DC deaktiv."

                        ${ELSE}
                               DetailPrint "Objektausführung in Adobe Reader DC war bereits deaktiv."
                        ${ENDIF}
        ClearErrors

        SectionEnd

SectionGroupEnd

Section "-Log"
             StrCpy $0 "$INSTDIR\install.log"
Push $0
Call DumpLog

SectionEnd


!macro getBefore Value Type Root Key Setting

       ReadReg${Type} ${Value} ${Root} "${Key}" "${Setting}"
        ClearErrors
!macroend


function .onInit
;	setShellVarContext all

	!insertmacro IsAdmin
	
  	;!insertmacro UnselectSection ${ScriptHost}
	!insertmacro getBefore $beforeScriptHost "DWORD" "HKCU" "SOFTWARE\Microsoft\Windows Script Host\Settings" "Enabled"
	!insertmacro getBefore $beforeAutoRun1 "DWORD" "HKCU" "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoDriveTypeAutoRun"
	!insertmacro getBefore $beforeUAC "DWORD" "HKLM" "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" "ConsentPromptBehaviorAdmin"
	!insertmacro getBefore $beforeShowFile "DWORD" "HKCU" "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt"
	
        ;!insertmacro UnselectSection "${PowerShell}"
        ;!insertmacro UnselectSection "${FileExtensions}"
        
        ReadRegDWORD $beforeAutorun1 HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoDriveTypeAutoRun"
        ReadRegDWORD $beforeAutorun2 HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoAutorun"
        ReadRegDWORD $beforeAutorun3 HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" "DisableAutoplay"

          StrCpy $officeVer "0"

        !insertmacro KeyExists "HKCU" "SOFTWARE\Microsoft\Office\" "12.0"
        Pop $R0

        ${IF} $R0 == 1
        StrCpy $officeVer "12.0";
                   
         ${ENDIF}

         !insertmacro KeyExists "HKCU" "SOFTWARE\Microsoft\Office\" "14.0"
          Pop $R0
              ${IF} $R0 == 1
         StrCpy $officeVer "14.0"

         ${ENDIF}
         
                 !insertmacro KeyExists "HKCU" "SOFTWARE\Microsoft\Office\" "15.0"
          Pop $R0
              ${IF} $R0 == 1
         StrCpy $officeVer "15.0"

         ${ENDIF}
     
         !insertmacro KeyExists "HKCU" "SOFTWARE\Microsoft\Office\" "16.0"
         Pop $R0
             ${IF} $R0 == 1
         StrCpy $officeVer "16.0"

         ${ENDIF}
      
         ${IF} $officeVer == "0"
      
         SectionSetFlags ${OLE} 16
	 SectionSetFlags ${ExcelMakros} 16
	 SectionSetFlags ${WordMakros} 16
	 SectionSetFlags ${PowerPointMakros} 16
	 SectionSetFlags ${ActiveX} 16
	 
	 ${ENDIF}

      
      
        ReadRegDWORD $beforeMakroWord HKCU "SOFTWARE\Microsoft\Office\$officeVer\Word\Security" "VBAWarnings"
        ReadRegDWORD $beforeMakroExcel HKCU "SOFTWARE\Microsoft\Office\$officeVer\Excel\Security" "VBAWarnings"
        ReadRegDWORD $beforeMakroPowerPoint HKCU "SOFTWARE\Microsoft\Office\$officeVer\PowerPoint\Security" "VBAWarnings"
        
        ReadRegDWORD $beforeOLEPowerPoint HKCU "SOFTWARE\Microsoft\Office\$officeVer\Excel\Security" "PackagerPrompt"
        ReadRegDWORD $beforeOLEPowerPoint HKCU "SOFTWARE\Microsoft\Office\$officeVer\Word\Security" "PackagerPrompt"
        ReadRegDWORD $beforeOLEPowerPoint HKCU "SOFTWARE\Microsoft\Office\$officeVer\PowerPoint\Security" "PackagerPrompt"
        ReadRegDWORD $beforeOfficeActiveX HKCU "SOFTWARE\Microsoft\Office\Common\Security" "DisableAllActiveX"

        ReadRegDWORD $beforeReaderDCX HKCU "SOFTWARE\Adobe\Acrobat Reader\DC\JSPrefs" "bEnableJS"
        ReadRegDWORD $beforeReaderXIX HKCU "SOFTWARE\Adobe\Acrobat Reader\XI\JSPrefs" "bEnableJS"
        ReadRegDWORD $beforeReaderXIObject HKCU "SOFTWARE\Adobe\Acrobat Reader\XI\Originals" "bAllowOpenFile"
        ReadRegDWORD $beforeReaderDCObject HKCU "SOFTWARE\Adobe\Acrobat Reader\DC\Originals" "bAllowOpenFile"
        
        
        ;Alle deaktivieren
        !insertmacro UnselectSection "${ScriptHost}"
        !insertmacro UnselectSection "${AutoPlay}"
        !insertmacro UnselectSection "${PowerShell}"
        !insertmacro UnselectSection "${UAC}"
        !insertmacro UnselectSection "${FileExtensions}"
        !insertmacro UnselectSection "${FileShowExtensions}"
        !insertmacro UnselectSection "${WordMakros}"
        !insertmacro UnselectSection "${ExcelMakros}"
        !insertmacro UnselectSection "${WordMakros}"
        !insertmacro UnselectSection "${PowerPointMakros}"

        !insertmacro UnselectSection "${OLE}"
        !insertmacro UnselectSection "${ActiveX}"
        !insertmacro UnselectSection "${ReaderJS}"
        !insertmacro UnselectSection "${ReaderObjects}"
        
        ;Auslesen, welche Maßnahmen schon aktiviert wurden:

        ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "WordOLE"

        ${If} ${Errors}
        
        ${Else}
        !insertmacro SelectSection "${OLE}"
        ${EndIf}
        
        ClearErrors

        ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "WordMakro"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${WordMakros}"
        ${EndIf}
        ClearErrors
        
        ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "ExcelMakro"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${ExcelMakros}"
        ${EndIf}
        ClearErrors
        
        ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "PowerPointMakro"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${PowerPointMakros}"
        ${EndIf}
        ClearErrors
        
           ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "OfficeX"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${ActiveX}"
        ${EndIf}
        ClearErrors
        
        
           ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "ReaderDCX"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${ReaderJS}"
        ${EndIf}
        ClearErrors
        
           ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "ReaderXIObject"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${ReaderObjects}"
        ${EndIf}
        ClearErrors
        
        ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "Scripthost"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${ScriptHost}"
        ${EndIf}
        ClearErrors

        ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "Autorun1"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${AutoPlay}"
        ${EndIf}
        ClearErrors

        ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "PowerShell"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${PowerShell}"
        ${EndIf}
        ClearErrors

        ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "UAC"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${UAC}"
        ${EndIf}
        ClearErrors

        ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "FileExtensions"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${FileExtensions}"
        ${EndIf}
        ClearErrors

           ReadRegDWORD $isset HKCU "SOFTWARE\CT Protector\" "ShowFile"

        ${If} ${Errors}

        ${Else}
                !insertmacro SelectSection "${FileShowExtensions}"
        ${EndIf}
        ClearErrors

          DeleteRegKey HKCU "SOFTWARE\CT Protector"

functionEnd

  ;Language strings
  LangString DESC_ScriptHost ${LANG_GERMAN} "Verhindert die Ausführung von VBScript und Javascript. Nicht geeignet für Entwickler."
  LangString DESC_AutoPlay ${LANG_GERMAN} "AutoRun und AutoPlay für alle Wechselmedien abschalten."
  LangString DESC_PowerShell ${LANG_GERMAN} "PowerShell und Eingabeaufforderung abschalten. Nicht für fortgeschrittene Benutzer."
  LangString DESC_UAC ${LANG_GERMAN} "User Account Control auf höchste Stufe einstellen. Sorgt für mehr Sicherheits-Abfragen."
  LangString DESC_FileExtensions ${LANG_GERMAN} "Zuordnung für Dateizuordnungen wie hta, js, jes,wsh, wsf, scr,vbs verhindern. Nicht für Entwickler geeignet."
  LangString DESC_FileShowExtensions ${LANG_GERMAN} "Dateiendungen immer anzeigen. Hat keine Nachteile."
 
  LangString DESC_WordMakros ${LANG_GERMAN} "Makros in Word abschalten. Nicht geeignet für Büro-Umgebungen."
  LangString DESC_ExcelMakros ${LANG_GERMAN} "Makros in Excel abschalten. Nicht geeignet für Büro-Umgebungen."
  LangString DESC_PowerPointMakros ${LANG_GERMAN} "Makros in PowerPoint abschalten. Nicht geeignet für Büro-Umgebungen."
  LangString DESC_OLE ${LANG_GERMAN} "OLE-Ausführung verhindern. Wird benutzt, um Office-Dokumente ineinander zu verschachteln."
  LangString DESC_ActiveX ${LANG_GERMAN} "ActiveX-Elemente in Office nicht ausführen. Hängt eng mit Makros zusammen, nicht für Büro-Umgebungen."

  LangString DESC_ReaderJS ${LANG_GERMAN} "In PDF-Dateien kann JavaScript integriert werden. Das bot in der Vergangenheit immer mal wieder Angriffsflächen, wird im Alltag zunehmend weniger eingesetzt. Interaktive Formulare funktionieren ohne JavaScript nicht mehr."
  LangString DESC_ReaderObjects ${LANG_GERMAN} "Die Einbettung von Objekten in PDF-Dateien wird nur sehr selten in der Praxis genutzt. Die negativen Folgen sind also gering."
  
  
  ;Assign language strings to sections
  !insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  
  !insertmacro MUI_DESCRIPTION_TEXT ${ScriptHost} $(DESC_ScriptHost)
  !insertmacro MUI_DESCRIPTION_TEXT ${AutoPlay} $(DESC_AutoPlay)
  !insertmacro MUI_DESCRIPTION_TEXT ${PowerShell} $(DESC_PowerShell)
  !insertmacro MUI_DESCRIPTION_TEXT ${UAC} $(DESC_UAC)
  !insertmacro MUI_DESCRIPTION_TEXT ${FileExtensions} $(DESC_FileExtensions)
  !insertmacro MUI_DESCRIPTION_TEXT ${FileShowExtensions} $(DESC_FileShowExtensions)

  !insertmacro MUI_DESCRIPTION_TEXT ${WordMakros} $(DESC_WordMakros)
  !insertmacro MUI_DESCRIPTION_TEXT ${ExcelMakros} $(DESC_ExcelMakros)
  !insertmacro MUI_DESCRIPTION_TEXT ${PowerPointMakros} $(DESC_PowerPointMakros)
  !insertmacro MUI_DESCRIPTION_TEXT ${OLE} $(DESC_OLE)
  !insertmacro MUI_DESCRIPTION_TEXT ${ActiveX} $(DESC_ActiveX)

  !insertmacro MUI_FUNCTION_DESCRIPTION_END

;--------------------------------
;Uninstaller Section

!insertmacro setDefaults "un."

Section "Uninstall"

  RMDir /r $INSTDIR
  RMDir /r "$SMPROGRAMS\c't Windows Protec'tor"
  DeleteRegKey HKLM "${UNINSTREG}"

  Call un.setDefaults
                
                DeleteRegKey HKCU "SOFTWARE\CT Protector"
                DetailPrint "Die Einstellungen wurden zurückgesetzt."

SectionEnd