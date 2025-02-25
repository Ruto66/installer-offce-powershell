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
                        |                                      0.1.1                                          |
                        |                              Code by Maxim Ruto66                                   |
                        |                                                                                     |
                        |_____________________________________________________________________________________|
"@
}

function Show-Menu {
    Write-Host "1. Install Office VISIO is Project. Office Full"
    Write-Host "2. Install Office don't VISIO is Project. Office NO Full"
    Write-Host "3. Install Office 2024 on Windows 11 Beta Version"
    Write-Host "4. Unistall to office"
    Write-Host "5. EXIT"
}

Show-Logo
Show-Menu

do {
    $choice = Read-Host "(1-5):"
} until ($choice -ge 1 -and $choice -le 5)


switch($choice){
    1{
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
    2{
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
    3{
        Write-Host "Office 2024 on Windows 11"
        function Show-Logo {
        Write-Host @"
                        _______________________________________________________________________________________ 
                        |                                                                                     |
                        |                                                                                     |
                        |                              Office 2024 on Windows 11                              |
                        |                                                                                     |
                        |                                                                                     |
                        |                                   Version  0.1.1                                    |
                        |                                Code by Maxim Ruto66                                 |
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
            $choice1 = Read-Host "(1-3):"
        } until ($choice1 -ge 1 -and $choice1 -le 3)
                switch($choice1){
                1{
                 # Определение XML структуры как строку
$xmlContent = @"
<Configuration ID="7a189364-5c40-4fe8-978a-86186a9a52d2">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB">
      <Language ID="ru-ru" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
    </Product>
    <Product ID="VisioPro2024Volume" PIDKEY="B7TN8-FJ8V3-7QYCP-HQPMV-YY89G">
      <Language ID="ru-ru" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
    </Product>
    <Product ID="ProjectPro2024Volume" PIDKEY="FQQ23-N4YCY-73HQ3-FM9WC-76HF4">
      <Language ID="ru-ru" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
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
<Configuration ID="80569198-7b5f-468f-bc8f-5bce496f196b">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB">
      <Language ID="ru-ru" />
      <ExcludeApp ID="Lync" />
    </Product>
    <Product ID="ProjectPro2024Volume" PIDKEY="FQQ23-N4YCY-73HQ3-FM9WC-76HF4">
      <Language ID="ru-ru" />
      <ExcludeApp ID="Lync" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
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
        Write-Host "Tolls"
        function Show-Logo {
        Write-Host @"
                        _______________________________________________________________________________________ 
                        |                                                                                     |
                        |                                                                                     |
                        |                                       Tolls                                         |
                        |                                                                                     |
                        |                                                                                     |
                        |                                   Version  0.1.1                                    |
                        |                                Code by Maxim Ruto66                                 |
                        |                                                                                     |
                        |_____________________________________________________________________________________|
"@
    }
            function Show-Menu {
            Write-Host "1. Unistall to 2007 - 2016"
            Write-Host "2. Unistall 2019,2021,2024"
            Write-Host "3. END"
        }
        Show-Logo
        Show-Menu
        do {
            $choice2 = Read-Host "(1-3):"
        } until ($choice2 -ge 1 -and $choice1 -le 3)
        switch($choice2){
            1{
                #Выполняем закрытие программ MS Office
                Stop-Process -Name OfficeClickToRun.exe.exe -Confirm
                Stop-Process -Name winword.exe -Confirm
                Stop-Process -Name excel.exe -Confirm
                # Указываем URL'а для загрузки файлов
                # Список URL для загрузки файлов
                $fileUrls = @(
                    "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub03.vbs"
                    "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub07.vbs"
                    "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub10.vbs"
                    "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub_O15msi.vbs"
                    "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub_O16msi.vbs"
                    "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrubc2r.vbs"
                    "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/Office2013Setup.exe"
                    "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/Office2016Setup.exe"
                    "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/Remove-PreviousOfficeInstalls.ps1"

                )

                # Укажите имя папки, которую вы хотите создать
                $folderName = "OfficeTolls"

                # Получите текущую директорию
                $currentDirectory = Get-Location

                # Создайте путь к новой папке
                $newFolderPath = Join-Path -Path $currentDirectory -ChildPath $folderName

                # Создайте новую папку, если она не существует
                if (-Not (Test-Path $newFolderPath)) {
                    New-Item -ItemType Directory -Path $newFolderPath
                }

                # Цикл для загрузки файлов
                foreach ($fileUrl in $fileUrls) {
                    # Укажите путь к сохранению загруженного файла
                    $filePath = Join-Path -Path $newFolderPath -ChildPath (Split-Path -Leaf $fileUrl)

                    # Скачайте файл
                    Invoke-WebRequest -Uri $fileUrl -OutFile $filePath

                    Write-Host "Файл загружен в $filePath"
                }
                Start-Process powershell.exe -ArgumentList $folderName\Remove-PreviousOfficeInstalls.ps1
                Remove-Item -Path $folderName
        
                Write-Host ""    
            }
            2{
                    # Указываем URL для загрузки файла
                    $url = "https://crystalidea.com/downloads/uninstalltool_portable.zip" # Замените на фактический URL файла
                    $fileName = "uninstalltool_portable.zip" # Имя файла, который будет загружен

                    # Получаем текущую директорию
                    $currentDir = Get-Location

                    # Путь к загружаемому файлу
                    $filePath = Join-Path -Path $currentDir -ChildPath $fileName

                    # Загружаем файл
                    Invoke-WebRequest -Uri $url -OutFile $filePath

                    # Проверяем, существует ли загруженный файл
                    if (Test-Path $filePath) {
                        # Указываем директорию для распаковки
                        $unzipFolder = Join-Path -Path $currentDir -ChildPath "Unzipped"

                        # Создаем директорию для распаковки, если она не существует
                        if (-not (Test-Path $unzipFolder)) {
                            New-Item -ItemType Directory -Path $unzipFolder
                        }

                        # Распаковываем файл
                        Expand-Archive -Path $filePath -DestinationPath $unzipFolder -Force

                        Write-Host "Файл загружен и распакован в папку: $unzipFolder"
                    } else {
                        Write-Host "Не удалось загрузить файл."
                    }
                    # Проверьте наличие файла
                    if (Test-Path "$unzipFolder\UninstallToolPortable.exe") {
                        # Сформируйте команду
                        $cmd = "$unzipFolder\UninstallToolPortable.exe"

                        # Запустите команду в CMD
                        Start-Process powershell.exe -ArgumentList $unzipFolder\UninstallToolPortable.exe -NoNewWindow -Wait
                        Remove-Item -Path $unzipFolder,$fileName
                    } else {
                        Write-Host "Файл UninstallToolPortable не найден по указанному пути."
                        Remove-Item -Path $unzipFolder,$fileName
                        Remove-Item -Path HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0 -Recurse -Force -ErrorAction Ignore
                        Remove-Item -Path HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0 -Recurse -Force -ErrorAction Ignore
                    }   
            }
            3{
                Write-Host "END"
            }
        }
            }
            5{
                Write-Host "END"
            }
}