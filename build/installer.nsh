; Allow NSIS to show the standard per-user / per-machine installation mode page.

!macro customInstall
  WriteRegStr HKCU "Software\Classes\Applications\OnyxPDF.exe" "" "OnyxPDF"
  WriteRegStr HKCU "Software\Classes\Applications\OnyxPDF.exe\DefaultIcon" "" "$INSTDIR\OnyxPDF.exe,0"
  WriteRegStr HKCU "Software\Classes\Applications\OnyxPDF.exe\shell\open\command" "" '"$INSTDIR\OnyxPDF.exe" "%1"'

  WriteRegStr HKCU "Software\Classes\OnyxPDF.PDF" "" "OnyxPDF PDF Document"
  WriteRegStr HKCU "Software\Classes\OnyxPDF.PDF\DefaultIcon" "" "$INSTDIR\OnyxPDF.exe,0"
  WriteRegStr HKCU "Software\Classes\OnyxPDF.PDF\shell\open\command" "" '"$INSTDIR\OnyxPDF.exe" "%1"'

  DeleteRegValue HKCU "Software\OnyxPDF\Capabilities" "ApplicationDescription"
  WriteRegStr HKCU "Software\OnyxPDF\Capabilities" "ApplicationName" "OnyxPDF"
  WriteRegStr HKCU "Software\OnyxPDF\Capabilities\FileAssociations" ".pdf" "OnyxPDF.PDF"
  WriteRegStr HKCU "Software\RegisteredApplications" "OnyxPDF" "Software\OnyxPDF\Capabilities"

  WriteRegStr HKCU "Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache" "$INSTDIR\OnyxPDF.exe.FriendlyAppName" "OnyxPDF"
  WriteRegStr HKCU "Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache" "$INSTDIR\OnyxPDF.exe.ApplicationCompany" "OnyxPDF Project"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, i 0, i 0)'
!macroend

!macro customUnInstall
  DeleteRegValue HKCU "Software\RegisteredApplications" "OnyxPDF"
  DeleteRegValue HKCU "Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache" "$INSTDIR\OnyxPDF.exe.FriendlyAppName"
  DeleteRegValue HKCU "Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache" "$INSTDIR\OnyxPDF.exe.ApplicationCompany"
  DeleteRegKey HKCU "Software\OnyxPDF\Capabilities\FileAssociations"
  DeleteRegKey /ifempty HKCU "Software\OnyxPDF\Capabilities"
  DeleteRegKey /ifempty HKCU "Software\OnyxPDF"
  DeleteRegKey HKCU "Software\Classes\OnyxPDF.PDF"
  DeleteRegKey HKCU "Software\Classes\Applications\OnyxPDF.exe"
!macroend
