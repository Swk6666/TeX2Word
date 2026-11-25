Unicode true
!include "MUI2.nsh"
!include "x64.nsh"
!include "LogicLib.nsh"
Name "TeX2Word"
OutFile "${OUTPUT}"
InstallDir "$LOCALAPPDATA\Programs\TeX2Word"
RequestExecutionLevel user
SetCompressor /SOLID /FINAL lzma
SetCompressorDictSize 32
AllowSkipFiles off
BrandingText "TeX2Word ${VERSION}"
VIProductVersion "${VERSION}.0"
VIAddVersionKey "ProductName" "TeX2Word"
VIAddVersionKey "ProductVersion" "${VERSION}"
VIAddVersionKey "FileVersion" "${VERSION}"
VIAddVersionKey "FileDescription" "TeX2Word 公式文档转换安装程序"
VIAddVersionKey "LegalCopyright" "TeX2Word"
!define MUI_ICON "${ICONFILE}"
!define MUI_UNICON "${ICONFILE}"
!define MUI_WELCOMEPAGE_TITLE "安装 TeX2Word"
!ifdef OFFLINE_MODELS
!define MUI_WELCOMEPAGE_TEXT "将中文、英文或中英混排的 PDF、TeX 和 Markdown 文档批量转换为 Word（原生可编辑 MathType 公式）。$\r$\n$\r$\n离线安装包已内置本地 PDF 识别模型，全程可离线使用。公式编辑使用电脑上已安装的 Microsoft Word 和桌面版 MathType 7。$\r$\n$\r$\n程序将安装在当前用户的个人应用目录中。"
!else
!define MUI_WELCOMEPAGE_TEXT "将中文、英文或中英混排的 PDF、TeX 和 Markdown 文档批量转换为 Word（原生可编辑 MathType 公式）。$\r$\n$\r$\nTeX 与 Markdown 转换安装后即可离线使用；首次转换 PDF 时会自动下载并缓存本地识别模型。公式编辑使用电脑上已安装的 Microsoft Word 和桌面版 MathType 7。$\r$\n$\r$\n程序将安装在当前用户的个人应用目录中。"
!endif
!define MUI_FINISHPAGE_RUN "$INSTDIR\TeX2Word.exe"
!define MUI_FINISHPAGE_RUN_TEXT "打开 TeX2Word"
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "SimpChinese"

!define UNINSTALLKEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\TeX2Word"

!macro CheckApp
  FindWindow $0 "" "TeX2Word · 公式文档转换"
  ${If} $0 != 0
    MessageBox MB_OK|MB_ICONEXCLAMATION "请关闭 TeX2Word 后重试。" /SD IDOK
    SetErrorLevel 2
    Abort
  ${EndIf}
!macroend
Function .onInit
  ${IfNot} ${RunningX64}
    MessageBox MB_ICONSTOP|MB_OK "请在 64 位 Windows 上安装 TeX2Word。" /SD IDOK
    SetErrorLevel 1
    Abort
  ${EndIf}
  !insertmacro CheckApp
FunctionEnd
Function un.onInit
  !insertmacro CheckApp
FunctionEnd

Section "TeX2Word"
  SetShellVarContext current
  SetRegView 64
  !include "${INSTALLFILES}"
  SetOutPath "$INSTDIR"
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  CreateShortCut "$SMPROGRAMS\TeX2Word.lnk" "$INSTDIR\TeX2Word.exe"
  CreateShortCut "$DESKTOP\TeX2Word.lnk" "$INSTDIR\TeX2Word.exe"
  WriteRegStr HKCU "${UNINSTALLKEY}" "DisplayName" "TeX2Word"
  WriteRegStr HKCU "${UNINSTALLKEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKCU "${UNINSTALLKEY}" "Publisher" "TeX2Word"
  WriteRegStr HKCU "${UNINSTALLKEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNINSTALLKEY}" "DisplayIcon" "$INSTDIR\TeX2Word.exe"
  WriteRegStr HKCU "${UNINSTALLKEY}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegStr HKCU "${UNINSTALLKEY}" "QuietUninstallString" '"$INSTDIR\Uninstall.exe" /S'
  WriteRegDWORD HKCU "${UNINSTALLKEY}" "EstimatedSize" ${SIZEKB}
  WriteRegDWORD HKCU "${UNINSTALLKEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINSTALLKEY}" "NoRepair" 1
  nsExec::ExecToLog '"$INSTDIR\runtime\python.exe" -I -B -m tex2word.winpath add "$INSTDIR"'
SectionEnd

Section "Uninstall"
  SetShellVarContext current
  SetRegView 64
  nsExec::ExecToLog '"$INSTDIR\runtime\python.exe" -I -B -m tex2word.winpath remove "$INSTDIR"'
  Delete "$SMPROGRAMS\TeX2Word.lnk"
  Delete "$DESKTOP\TeX2Word.lnk"
  !include "${REMOVEFILES}"
  Delete "$INSTDIR\Uninstall.exe"
  DeleteRegKey HKCU "${UNINSTALLKEY}"
  SetOutPath "$TEMP"
  RMDir "$INSTDIR"
SectionEnd
