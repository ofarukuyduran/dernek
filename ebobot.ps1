if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -Command `"cd '$pwd'; & '$PSCommandPath';`"";
    exit;
}
try {
    Write-Host ""

    Write-Host "########################################################################################" -ForegroundColor Yellow

    Write-Host ""

    Write-Host "Windows icin EbaBot Katılımsız Yükleme Araci v0.1" -ForegroundColor Yellow

    Write-Host ""

    Write-Host "Bu yükleme aracı ebabotun kurulumunu bir kaç dk. içinde katılımsız bir şekilde yaparak "  -ForegroundColor Yellow
    Write-Host "Eren Mustafa ÖZDAL'ın ebabot projesine katkıda bulunmak amacıyla yazılmıştır." -ForegroundColor Yellow

    Write-Host ""

    Write-Host "Desteklenen isletim sistemleri:

          Windows 10 (32bit/64 bit)" -ForegroundColor Yellow

    Write-Host ""

    Write-Host "Hazirlayan: Ömer Faruk UYDURAN | Geri bildirim: farukuyduran@gmail.com" -ForegroundColor Yellow

    Write-Host ""

    Write-Host "#######################################################################################" -ForegroundColor Yellow

    Write-Host ""


    function Check_Program_Installed {
        [CmdletBinding()]
        Param(
            [Parameter(Position = 0, Mandatory=$true, ValueFromPipeline = $true)]
            $Name
        )
        $app = Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' | 
                    #Where-Object { $_.DisplayName -eq $Name } |
                    Where-Object { $_.DisplayName -match $Name } | 
                    Select-Object DisplayName, DisplayVersion, InstallDate, Version
        if ($app) {
            return $app.DisplayVersion
        } 
    }

    function Set-ShortCut {
    param ( [string]$SourceExe, [string]$DestinationPath )

    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($DestinationPath)
    $Shortcut.TargetPath = $SourceExe
    $Shortcut.WorkingDirectory = "$env:USERPROFILE\Downloads\ebabot-0.3\"
    $Shortcut.Save()
    }
      
    function DownloadFile($url, $targetFile)

    {

       $uri = New-Object "System.Uri" "$url"
       $request = [System.Net.HttpWebRequest]::Create($uri)
       $request.set_Timeout(15000) #15 second timeout
       $response = $request.GetResponse()
       $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)
       $responseStream = $response.GetResponseStream()
       $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create
       $buffer = new-object byte[] 10KB
       $count = $responseStream.Read($buffer,0,$buffer.length)
       $downloadedBytes = $count
       while ($count -gt 0)
       {
           $targetStream.Write($buffer, 0, $count)
           $count = $responseStream.Read($buffer,0,$buffer.length)
           $downloadedBytes = $downloadedBytes + $count

           Write-Progress -activity "Dosya İndiriliyor '$($url.split('/') | Select -Last 1)'" -status "İndirilen ($([System.Math]::Floor($downloadedBytes/1024))K of $($totalLength)K): " -PercentComplete ((([System.Math]::Floor($downloadedBytes/1024)) / $totalLength)  * 100)
       }

         Write-Progress -activity "Dosyanın indirilmesi tamamlandı '$($url.split('/') | Select -Last 1)'"
     
       $targetStream.Flush()
       $targetStream.Close()
       $targetStream.Dispose()
       $responseStream.Dispose()
    }

    if (Check_Program_Installed "Python"){
        Write-Host "Python sisteminizde zaten kurulu."  -ForegroundColor Green
    } else {
        Write-Host "Python sisteminzide kurulu değil indirilip kurulacak."  -ForegroundColor Green
        if (-not (Test-Path "C:\tmp\")) {
                New-Item -ItemType directory -Path "C:\tmp\" 
                Write-Host "C:\tmp\ dizini oluşturuldu."  -ForegroundColor Green
            }
             Write-Host "Python İndiriliyor...Bekleyiniz"  -ForegroundColor Yellow
            #Invoke-WebRequest https://www.python.org/ftp/python/3.9.1/python-3.9.1-amd64.exe -OutFile C:\tmp\python-3.9.1-amd64.exe
            downloadFile "https://www.python.org/ftp/python/3.9.2/python-3.9.2.exe" "C:\tmp\python-3.9.2.exe"
            Write-Host "Python İndirildi"  -ForegroundColor Green

            if( Test-Path "C:\tmp\python-3.9.2.exe"){
                Write-Host "İndirilen Python Kurulum dosyası kontrol edildi."  -ForegroundColor Green
            } else {
                Write-host "Python kurulum dosyası bulunamadı. Tekrar indirilecektir."  -ForegroundColor Red
                downloadFile "https://www.python.org/ftp/python/3.9.2/python-3.9.2.exe" "C:\tmp\python-3.9.2.exe"
            }
            Write-Host "Python Kuruluyor. Lütfen Bekleyiniz..."  -ForegroundColor Yellow

            $process = Start-Process -FilePath "c:\tmp\python-3.9.2.exe" -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -wait

    for($i = 0; $i -le 100; $i++ % 100)
    {
        Write-Progress -Activity "Python Kuruluyor" -PercentComplete $i -Status "Kurulan"
        Start-Sleep -Milliseconds 100
        if ($process.HasExited) {
            Write-Progress -Activity "Kuruluyor" -Completed
             break
        } 
    }
        Write-Host "Python Kuruldu."  -ForegroundColor Green
         }

    if (Check_Program_Installed "Google Chrome"){
         Write-Host "Google Chrome sisteminizde zaten kurulu."  -ForegroundColor Green
    } else {
         Write-Host "Google Chrome sisteminizde kurulu değil.Siteminize Google Chrome indirilip kurulacak."  -ForegroundColor Yellow
         if (-not (Test-Path "C:\tmp\")) {
                New-Item -ItemType directory -Path "C:\tmp\" 
                Write-Host "C:\tmp\ dizini oluşturuldu."  -ForegroundColor Green
            }
            Write-Host "Google Chrome İndiriliyor...Bekleyiniz" 
            #Invoke-WebRequest https://www.python.org/ftp/python/3.9.1/python-3.9.1-amd64.exe -OutFile C:\tmp\python-3.9.1-amd64.exe
            downloadFile " http://dl.google.com/chrome/install/latest/chrome_installer.exe" "c:\tmp\chrome_installer.exe"
            Write-Host "Google Chrome İndirildi"  -ForegroundColor Green

             if( Test-Path "C:\tmp\chrome_installer.exe"){
                Write-Host "İndirilen Google Chrome kontrol edildi."  -ForegroundColor Green
            } else {
                Write-host "Google Chrome dosyası bulunamadı. Tekrar indirilecektir."  -ForegroundColor Red
                downloadFile " http://dl.google.com/chrome/install/latest/chrome_installer.exe" "c:\tmp\chrome_installer.exe"
            }
            Write-Host "Google Chrome Kuruluyor. Lütfen Bekleyiniz..." -ForegroundColor Yellow
            $process = Start-Process -FilePath "c:\tmp\chrome_installer.exe" -ArgumentList "/silent /install" -wait

            for($i = 0; $i -le 100; $i++ % 100)
            {
                Write-Progress -Activity "Google Chrome Kuruluyor" -PercentComplete $i -Status "Kurulan"
                Start-Sleep -Milliseconds 100
                if ($process.HasExited) {
                    Write-Progress -Activity "Kuruluyor" -Completed
                     break
                } 
            }

            Write-Host "Google Chrome Kuruluyor. Lütfen Bekleyiniz..." -ForegroundColor Green

    }

        Write-Host "Ebabot indiriliyor . Lütfen Bekleyiniz..."  -ForegroundColor Yellow
        if (-not (Test-Path "C:\tmp\")) {
                New-Item -ItemType directory -Path "C:\tmp\" 
                Write-Host "C:\tmp\ dizini oluşturuldu."  -ForegroundColor Green
            }
            #downloadFile "https://github.com/erenmustafaozdal/ebabot/archive/v0.3.zip" "c:\tmp\v0.3.zip"
            Invoke-WebRequest https://github.com/erenmustafaozdal/ebabot/archive/v0.3.zip -OutFile C:\tmp\v0.3.zip
            Write-Host "Ebabot indirildi"  -ForegroundColor Green
         if( -not (Test-Path "C:\tmp\v0.3.zip")){
                Write-host "Ebabot bulunamadı. Tekrar indirilecektir."  -ForegroundColor Red
                Invoke-WebRequest https://github.com/erenmustafaozdal/ebabot/archive/v0.3.zip -OutFile C:\tmp\v0.3.zip
            }

            Write-Host "Ebabot zip'ten çıkarılacak"  -ForegroundColor Yellow

    Get-ChildItem 'C:\tmp\v0.3.zip'  | Expand-Archive -DestinationPath $env:USERPROFILE\Downloads -Force
    Write-Host "Ebabot $env:USERPROFILE\Downloads\ebabot-0.3\ dizinine çıkarıldı"  -ForegroundColor Green
    Write-Host "Ebabot'un çalışması için Chrome driver yüklenecek"  -ForegroundColor Yellow
    Write-Host "Google Chrome Driver indiriliyor..."  -ForegroundColor Yellow
    #Invoke-WebRequest https://chromedriver.storage.googleapis.com/89.0.4389.23/chromedriver_win32.zip -OutFile C:\tmp\chromedriver_win32.zip
    downloadFile "https://chromedriver.storage.googleapis.com/89.0.4389.23/chromedriver_win32.zip" "c:\tmp\chromedriver_win32.zip"
    Write-Host "Google Chrome Driver indirildi. Zip'ten çıkarılacak"  -ForegroundColor Green
    Get-ChildItem 'C:\tmp\chromedriver_win32.zip'  | Expand-Archive -DestinationPath $env:USERPROFILE\Downloads\ebabot-0.3 -Force
    Write-Host "Google Chrome Driver $env:USERPROFILE\Downloads\ebabot-0.3\ dizinine çıkarıldı"  -ForegroundColor Green
    cd $env:USERPROFILE\Downloads\ebabot-0.3
    pip install -r requirements.txt 
    Write-Host "env dosyası oluşturalacak"  -ForegroundColor Yellow
    $envfile = "$env:USERPROFILE\Downloads\ebabot-0.3\.env"
    Set-Content -encoding "UTF8" $envfile ('DRIVER_PATH="' +"$env:USERPROFILE\Downloads\ebabot-0.3\chromedriver.exe"+'"'+"`r`n"+

    'USERS_EXCEL="'+"$env:USERPROFILE\Downloads\ebabot-0.3\excel templates\örnek kullanıcılar.xls"+'"'+"`r`n"+"`r`n"+

    'WEB_HEADLESS=False
    WEB_IMPLICITLY_WAIT=3
    EBA_USER_LOGIN=False
    WEB_SIZE="max"
    # WEB_SIZE="1920,1080"')
    Write-Host "env dosyası oluşturuldu"  -ForegroundColor Green

    Set-ShortCut "$env:USERPROFILE\Downloads\ebabot-0.3\main.py" "$Home\Desktop\Ebabot.lnk"
    Write-Host "Ebabot'un kısayolu masaüstüne oluşturuldu"  -ForegroundColor Green
    Set-ShortCut "$env:USERPROFILE\Downloads\ebabot-0.3\excel templates\örnek kullanıcılar.xls" "$Home\Desktop\Kullanıcılar.lnk"
    Write-Host "Excel Kullanıcı dosyası kısayolu masaüstüne oluşturuldu"  -ForegroundColor Green
    Set-ShortCut "$env:USERPROFILE\Downloads\ebabot-0.3\excel templates\örnek dersler.xls" "$Home\Desktop\Dersler.lnk"
    Write-Host "Excel Kullanıcı dosyası kısayolu masaüstüne oluşturuldu
    Ebabot kurulumu bitti.Masaüstündeki ebabot kısayoluna çift tılayarak çalıştırabilirsiz"  -ForegroundColor Green
    Start-Sleep -Second 5
}
catch{
    $exception = $_.Exception.Message
    Out-File -FilePath 'c:\tmp\ebabot.log' -Append -InputObject $exception
	
}
