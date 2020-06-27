!pragma warning error all
!include "MUI2.nsh"

Name "EliBackup v1.3"
OutFile "EliInstaller(EliBackup v1.3).exe"
Unicode True
InstallDir "$PROGRAMFILES"
RequestExecutionLevel admin

!define MUI_LANGDLL_ALLLANGUAGES
!define MUI_LANGDLL_REGISTRY_ROOT "HKCU" 
!define MUI_LANGDLL_REGISTRY_KEY "Software\Modern UI Test" 
!define MUI_LANGDLL_REGISTRY_VALUENAME "Installer Language"

!define MUI_ICON "D:\Desktop\EliServices\EliBackup\Logo.ico"
!define MUI_UNICON "D:\Desktop\EliServices\EliBackup\Logo.ico"
!define MUI_BGCOLOR 1FEC28
!define MUI_FINISHPAGE_NOAUTOCLOSE

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "${NSISDIR}\Docs\Modern UI\License.txt"
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_LANGUAGE "German"
!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_LANGUAGE "French"
!insertmacro MUI_LANGUAGE "Spanish"
!insertmacro MUI_LANGUAGE "Italian"
!insertmacro MUI_LANGUAGE "Dutch"

!insertmacro MUI_RESERVEFILE_LANGDLL

Section "EliBackup Basis Installation" SecBasic
  
  CreateDirectory "$INSTDIR\EliServices"
  CreateDirectory "$INSTDIR\EliServices\EliBackup"
  SetOutPath "$INSTDIR\EliServices\EliBackup"
  File "D:\Desktop\EliServices\EliBackup\Versionen\1\1.3\elibackup.config"
  File "D:\Desktop\EliServices\EliBackup\Versionen\1\1.3\elibackup.vbs"
  File "D:\Desktop\EliServices\EliBackup\Versionen\1\1.3\Readme.txt"

SectionEnd

Section "Release Notes" SecRelNotes

  SetOutPath "$INSTDIR\EliServices\EliBackup"
  File "D:\Desktop\EliServices\EliBackup\Versionen\1\1.3\Release Notes v1.3.txt"

SectionEnd

Section "Icons" SecIco

  SetOutPath "$INSTDIR\EliServices\EliBackup"
  File "D:\Desktop\EliServices\EliBackup\Versionen\1\1.3\EliBackup-Icon.ico"
  File "D:\Desktop\EliServices\EliBackup\Versionen\1\1.3\Logo.ico"

SectionEnd

Function .onInit

  !insertmacro MUI_LANGDLL_DISPLAY

FunctionEnd