"Import-Module .\InstallerOffice.ps1"
# Определение XML структуры как строку
$xmlContent = @"
<Configuration ID="041092c8-0ee2-4afc-8c73-f1102be4e8a0">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume" PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Publisher" />
    </Product>
    <Product ID="VisioPro2021Volume" PIDKEY="KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Publisher" />
    </Product>
    <Product ID="ProjectPro2021Volume" PIDKEY="FTNWT-C6WBT-8HMGF-K9PRX-QV9H8">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Publisher" />
    </Product>
    <Product ID="LanguagePack">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="60" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="52" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="ODT" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
</Configuration>
"@

# Преобразование XML строки в объект XML
$xmlObject = [xml]$xmlContent
#########################################
$xmlContent1 = @"
<Configuration ID="1abeb455-e28a-40a2-b72f-6abac17aef95">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume" PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
      <ExcludeApp ID="Lync" />
    </Product>
    <Product ID="LanguagePack">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="60" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="52" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="ODT" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
</Configuration>
"@





function Show-Logo {
    Write-Host @"
                        _______________________________________________________________________________________ 
                        |                                                                                     |
                        |                                                                                     |
                        |                             Office Installe's GUI                                   |
                        |                                                                                     |
                        |                                                                                     |
                        |                                                                                     |
                        |                                                                                     |
                        |                                                                                     |
                        |_____________________________________________________________________________________|
"@
}

function Show-Menu {
    Write-Host "1. Install Office VISIO is Project. Office Full"
    Write-Host "2. Install Office don't VISIO is Project. Office NO Full"
    Write-Host "3. Install Office 2024 on Windows 11"
    Write-Host "4. EXIT"
}

Show-Logo
Show-Menu

do {
    $choice = Read-Host "(1-4):"
} until ($choice -ge 1 -and $choice -le 4)

switch ($choice) {
    1 {
        $currentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
        Write-Host "Install to go"
        # Путь к файлу, который нужно создать
        $filePath = "$currentDirectory\ConfigurationOffice2021Full.xml"
        # Создаем XML-документ и записываем его в файл
        $xmlDocument = New-Object System.Xml.XmlDocument
        $xmlDocument.LoadXml($xmlContent)
        $xmlDocument.Save($filePath)
        Write-Host "XML файл успешно создан в $filePath"
        $oldGeoID=(Get-WinHomeLocation).GeoId
        Set-WinHomeLocation -GeoId 244
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Experiment -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentEcs -Recurse -Force -ErrorAction Ignore
        Start-Process -FilePath $currentDirectory\Setup.exe -ArgumentList "/configure",$currentDirectory\ConfigurationOffice2021Full.xml # Вывод объекта XML
        Start-Sleep -Seconds 60
        Remove-Item -Path $filePath
        Write-Host "XML файл успешно удалён в $filePath"
    }
    2 {
        $currentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
        Write-Host "Install to go"
        # Путь к файлу, который нужно создать
        $filePath = "$currentDirectory\ConfigurationOffice2021_NO_VISIO_Project.xml"
        # Создаем XML-документ и записываем его в файл
        $xmlDocument = New-Object System.Xml.XmlDocument
        $xmlDocument.LoadXml($xmlContent1)
        $xmlDocument.Save($filePath)
        Write-Host "XML файл успешно создан в $filePath"
        $oldGeoID=(Get-WinHomeLocation).GeoId
        Set-WinHomeLocation -GeoId 244
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Experiment -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentEcs -Recurse -Force -ErrorAction Ignore
        Start-Process -FilePath $currentDirectory\Setup.exe -ArgumentList "/configure",$currentDirectory\ConfigurationOffice2021_NO_VISIO_Project.xml # Вывод объекта XML
        Start-Sleep -Seconds 60
        Remove-Item -Path $filePath
        Write-Host "XML файл успешно удалён в $filePath"
    }
    3 {
        Write-Host "Office 2024 on Windows 11"
        function Show-Logo {
        Write-Host @"
                        _______________________________________________________________________________________ 
                        |                                                                                     |
                        |                                                                                     |
                        |                              Office 2024 on Windows 11                              |
                        |                                                                                     |
                        |                                                                                     |
                        |                                       Clouse                                        |
                        |                                                                                     |
                        |                                                                                     |
                        |_____________________________________________________________________________________|
"@
    }
    function Show-Menu {
    Write-Host "1. Office 2024 Full"
    Write-Host "2. Office 2024"
    Write-Host "3. END"
}
Show-Logo
Show-Menu
do {
    $choice1 = Read-Host "(1-2):"
} until ($choice1 -ge 1 -and $choice1 -le 2)
switch ($choice1){
        1{
            # Определение XML структуры как строку
$xmlContent = @"
<Configuration ID="041092c8-0ee2-4afc-8c73-f1102be4e8a0">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume" PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Publisher" />
    </Product>
    <Product ID="VisioPro2024Volume" PIDKEY="KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Publisher" />
    </Product>
    <Product ID="ProjectPro2024Volume" PIDKEY="FTNWT-C6WBT-8HMGF-K9PRX-QV9H8">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Publisher" />
    </Product>
    <Product ID="LanguagePack">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="60" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="52" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="ODT" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
</Configuration>
"@
        $currentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
        Write-Host "Install to go"
        # Путь к файлу, который нужно создать
        $filePath = "$currentDirectory\ConfigurationOffice2024Full.xml"
        # Создаем XML-документ и записываем его в файл
        $xmlDocument = New-Object System.Xml.XmlDocument
        $xmlDocument.LoadXml($xmlContent)
        $xmlDocument.Save($filePath)
        Write-Host "XML файл успешно создан в $filePath"
        $oldGeoID=(Get-WinHomeLocation).GeoId
        Set-WinHomeLocation -GeoId 244
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Experiment -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentEcs -Recurse -Force -ErrorAction Ignore
        Start-Process -FilePath $currentDirectory\Setup.exe -ArgumentList "/configure",$currentDirectory\ConfigurationOffice2024Full.xml # Вывод объекта XML
        Start-Sleep -Seconds 120
        Remove-Item -Path $filePath
        Write-Host "XML файл успешно удалён в $filePath" 
        }
        2{
# Преобразование XML строки в объект XML
$xmlObject = [xml]$xmlContent
#########################################
$xmlContent1 = @"
<Configuration ID="1abeb455-e28a-40a2-b72f-6abac17aef95">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume" PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
      <ExcludeApp ID="Lync" />
    </Product>
    <Product ID="LanguagePack">
      <Language ID="ru-ru" />
      <Language ID="en-us" />
      <Language ID="MatchPreviousMSI" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="60" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="52" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="ODT" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
</Configuration>
"@
        $currentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
        Write-Host "Install to go"
        # Путь к файлу, который нужно создать
        $filePath = "$currentDirectory\ConfigurationOffice2024.xml"
        # Создаем XML-документ и записываем его в файл
        $xmlDocument = New-Object System.Xml.XmlDocument
        $xmlDocument.LoadXml($xmlContent1)
        $xmlDocument.Save($filePath)
        Write-Host "XML файл успешно создан в $filePath"
        $oldGeoID=(Get-WinHomeLocation).GeoId
        Set-WinHomeLocation -GeoId 244
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Experiment -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentEcs -Recurse -Force -ErrorAction Ignore
        Start-Process -FilePath $currentDirectory\Setup.exe -ArgumentList "/configure",$currentDirectory\ConfigurationOffice2024.xml # Вывод объекта XML
        Start-Sleep -Seconds 60
        Remove-Item -Path $filePath
        Write-Host "XML файл успешно удалён в $filePath"

        }
        3{
            Write-Host "END"
        }
}
    }
    4{
        Write-Host "END"
    }
}

