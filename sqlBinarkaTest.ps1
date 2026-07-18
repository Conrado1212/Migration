$LogDir  = ".\logs"
$LogFile = Join-Path $LogDir "log_$(Get-Date -Format 'yyyy-MM-dd').txt"


if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}
function Write-MessageLog{
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO",
        [ConsoleColor]$Color = "White"
    )

    $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    $entry = "$timestamp - $Level - $Message"

    switch ($Level) {
        "INFO"    { Write-Host $entry -ForegroundColor $Color }
        "WARNING" { Write-Warning $Message }
        "ERROR"   { Write-Error $Message }
    }

    Add-Content -Path $LogFile -Value $entry
}

Write-MessageLog "This PowerShell script will migrate your data from xlsx or table to System" 
 
 
 
 
 #$CLIENT_ID = Get-Secret -Name 'CLIENT_ID' -AsPLainText


#$CLIENT_SECRET = Get-Secret -Name 'CLIENT_SECRET' -AsPLainText

$baseUrl = ""
 
if(-not $baseUrl){
    $baseUrl = Read-Host "[Line 9] Write your baseUrl (example. https://example.com)"
   }
 function Get-DataFromExcel{
    param(
        [string]$path
    )
 $excel = New-Object -ComObject Excel.Application
   # if([System.IO.Path]::GetExtension($path) -ne ".xlsx"){
   #     Write-Host "The file must have the xlsx extension"
   #     exit 1
  #  }
        $workbook = $excel.Workbooks.Open("")
        $sheet = $workbook.Sheets.Item(1)
        $dane =@()
         $row = 1
         while($sheet.Cells.Item($row, 1).Text -ne ""){
              if($sheet.Cells.Item($row, 5).Text -eq "AKTYWNA"){
            try{
                    # --- RESET zmiennych plikowych ---
                    $att =@()

                   #--------------------------------------
           
                  $bpsId = $sheet.Cells.Item($row,1).Text
                  $DataZawarciaUmowy = $sheet.Cells.Item($row,2).Text
                  $DataObowUmowy = $sheet.Cells.Item($row,3).Text
                  $umowaLata = $sheet.Cells.Item($row,9).Text
                  $kontrahent = $sheet.Cells.Item($row,6).Text
                  $projekt = $sheet.Cells.Item($row,7).Text
                  $Okresochrony = $sheet.Cells.Item($row,8).Text
                 $paths = @( 
                     $sheet.Cells.Item($row,11).Text
                     $sheet.Cells.Item($row,12).Text
                 ) 
                   
                foreach($filePath in $paths){
                    if ([string]::IsNullOrWhiteSpace($filePath) ) {    
                        Write-MessageLog "Sciezka jest pusta $row" -Level "WARNING"
                        continue
                     }

                     if(-not (Test-Path  $filePath)){
                        Write-MessageLog "Plik nie istnieje: $filePath wiersz $row" -Level "WARNING"
                        continue
                     }
                     $bytes = [System.IO.File]::ReadAllBytes($filePath)
                        #nazwa pliku ze sciezki 
                        $name = Split-Path -Path  $filePath -Leaf
                        Write-MessageLog "Nazwa pliku: $name" -Level "INFO"

                        $att +=[PSCustomObject]@{
                            Name = $name
                            Bytes = $bytes
                        }
                }
              
                    if(-not ($dane | Where-Object {$_.bpsId -eq $bpsId})){
                    $dane += [PSCustomObject]@{
                             bpsId = $bpsId
                             DataZawarciaUmowy = $DataZawarciaUmowy
                             DataObowiazywaniaUmowy = $DataObowUmowy
                             firma = $kontrahent
                             cel = $projekt
                             okresOchrony = $Okresochrony
                             umowaLata = $umowaLata
                             Zalaczniki = $att
                             wfd_id = $wfd_id
                    }
                }
            }catch{
                Write-MessageLog "[Line 110] E in row $row : $_" -Level "WARNING"
            }
        }
            $row++;
        }

        $workbook.Close($false)
        $excel.Quit()
        return  $dane 
    }
#zmienic na static i elo 
   # $path = Read-Host "Please provide the path to the Excel file"
  # $path = "C:\Users\k.krawczyk\Desktop\xd\fizyk\powershellScript\Book1.xlsx"
#if([string]::IsNullOrWhiteSpace($path)){
  #  Write-Host "the Path variable cannot be empty"
  #  exit 1
#}else{
   $result =  Get-DataFromExcel -path $path
   foreach ($item in $result) {
    Write-MessageLog ("Data item: $($item.bpsId)" + $item)
    foreach($att in $item.Zalaczniki){
        Write-MessageLog "Data att $($item.bpsId): $($att.Name)"
    }
}
#}


function Get-AccessToken{
    #parametr base url ktory podaje user
    #pobranei clientId oraz cleitnsecret jako zmiennej srodwoiskowej 
        param(
            [string]$baseUrl,
            [string]$clientId,
            [string]$clientSecret
        )
        if ([string]::IsNullOrWhiteSpace($clientId)) {

        $clientId = Read-Host "[Line 26] Write  CLIENT_ID"
        }else{
            Write-MessageLog  "[Line 28] Variable clientId value retrieved"
        }

        if ([string]::IsNullOrWhiteSpace($clientSecret)) {
             $clientSecret = Read-Host "Write CLIENT_SECRET"
        }else{
            Write-MessageLog   "[Line 33] Variable clientSecret value retrieved"
        }
       
        
        #url metody uwierzytelnienia 
        $tokenUrl = "$baseUrl/api/oauth2/token"
        #przekazanei body
        $body =@{
            grant_type = "client_credentials"
            client_id = $clientId
            client_secret = $clientSecret
        }
        try{
            #proba otrzymanai accesstoekntu za pomoca wylownaia invoke-RestMethod 
            $respone = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded" 
            return $respone.access_token
        }catch{
            #obsluga bledu 
            Write-MessageLog "[Line 51] Error authentication: $_" -Level "ERROR"
            return $null
        }
} 




function Start-New{
    param(
        [string]$baseUrl,
        [string]$clientId,
        [string]$clientSecret
    )
    $token = Get-AccessToken -baseUrl $baseUrl -clientId $clientId -clientSecret $clientSecret

    if (-not $token) {
        Write-MessageLog "[Line 87] No token" -Level "ERROR"
        return
    }

    $headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
    }
    $new = "$baseUrl/api/data/v6.0/db/1/elements?path=ef583612-fbc2-4705-a239-8049956e929e"
    
   $dane =  Get-DataFromExcel 
 
   $batchSize = 10

   for ($i = 0; $i -lt $dane.Count; $i += $batchSize) {

       }
       #$batch = $dane[$i..([Math]::Min($i + $batchSize - 1, $dane.Count - 1))]
       #$dane | ConvertTo-Json | Set-Content "stan_$(Get-Date -Format 'yyyy-MM-dd').json"
#check in 

#status = "new" / "created" / "updated" / "error"
    foreach($item in $dane){
        
        $body =@{
            workflow = @{
                guid = "5e0ebcc9-8cfa-4384-a87e-7fdceb3a249e"
            }
            formType = @{
                guid = "95f176f7-9c7d-4877-8263-61a571d016c0"
            }
            formFields = @(
                @{
                     guid = "b7a98539-80c8-497b-a873-b6b169aa8bb6"   #1Data wejścia umowy w życie
                    svalue = "$($item.col2)"
                },
                @{
                     guid = "c426d8f1-3b6d-4695-8f8d-f77aafccef98" #  2 [tech] umowa + lata
                    svalue = "$($item.col2)"
                },
                @{
                     guid = "4afa9d35-ce1c-445f-9bcc-a85f8b490070" #  3 Czas obowiązywania umowy (lata)
                     svalue = "$($item.col3)"
                },
                @{
                    guid = "3da30053-1f5c-402a-9833-a781baac1ecc" #  5Nazwa firmy
                     svalue = "$($item.col5)"
                },
                @{
                     guid = "d37bc66d-5d6f-4a53-976a-2994df2424fa" #  6Cel i zakres
                      svalue = "$($item.col8)"
                },
                @{
                     guid = "5c49bbbe-d062-4202-975c-b3ac360e3a8d" #  7Okres ochrony (lata)
                    svalue = "$($item.col9)"
                }
            )
            #zmienic i testy elo 
            attachments = $item.Zalaczniki

           } | ConvertTo-Json -Depth 5
          # $allGroupsParam +=$body
          $testBody = [System.Text.Encoding]::UTF8.GetBytes($body)
          
          # Write-Host "[Line 121] New Instance Params $Body"
         # Write-MessageLog "[Line 121] New Instance Params $testBody"
            try{
          # $response =  Invoke-RestMethod -Method Post -Uri $new -Body $Body  -Headers $headers 
           $response =  Invoke-RestMethod -Method Post -Uri $new -Body $testBody  -Headers $headers 
           $item.wfd_id = $response.result.wfdId
           Write-MessageLog "[Line 124] Instance created for " $item.id $response.result.wfdId
           Write-MessageLog "[Line 125] Response: $($response | ConvertTo-Json -Depth 5)"
        }catch{
            Write-MessageLog "[Line 126] Error : $_" -Level "ERROR"
        return $null
    }
    }  
        # $response | ConvertTo-Json -Depth 5
}
#$response = @(Start-New -baseUrl $baseUrl -clientId $CLIENT_ID -clientSecret $CLIENT_SECRET)

#Write-Host "[line 134] data: $response"