$LogDir  = "C:\logsPower\logs"
$LogFile = Join-Path $LogDir "log_{WFD_ID}_$(Get-Date -Format 'yyyy-MM-dd').txt"
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

$baseUrl = ""
 $CLIENT_ID = ""
$CLIENT_SECRET = ""
function Get-AccessToken{
param(
            [string]$baseUrl,
            [string]$clientId,
            [string]$clientSecret
        )
$tokenUrl = "$baseUrl/api/oauth2/token"
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


function Get-File{
    $att =@()
    
    $testFiles = @(
        @{  
            Value = "{COLUMN_Value:DET_Att6#DET_Att6}"
            Separator = $null
        },
        @{  
            Value =  "{COLUMN_Value:det_att7#det_att7}"
            Separator =  "(?=C:\\certs\\)"
        } 
   );

    foreach ($testFile in $testFiles) {
        if ([string]::IsNullOrWhiteSpace($testFile)) {
            continue
        }

        if($testFile.Separator){
            $files = $testFile.Value -split $testFile.Separator |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ }
        }else{
            $files = @($testFile.Value.Trim())
        }
        foreach ($file in $files) {
if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
$msg = "Plik nie istnieje: $file"
Write-MessageLog $msg -Level "ERROR"
throw [System.IO.FileNotFoundException]::new($msg)
}
try{
     $bytes = [System.IO.File]::ReadAllBytes($file)
      $name = Split-Path -Path  $testFile -Leaf
       $att +=[PSCustomObject]@{
                            name = $name
                            content = [Convert]::ToBase64String($bytes)
                        }
}catch{
$msg2 = "Nie udało się odczytać pliku '$file'. Błąd: $($_.Exception.Message)"
Write-MessageLog $msg2 -Level "ERROR"
throw $msg
}
        }
                    }
                        return $att
}
#$file = Get-File
#Write-Output $file
function add-Attachemts{
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
    $newAtt = "$baseUrl/api/data/v6.0/db/1/elements/{COLUMN_Value:DET_Att5#DET_Att5}?forceCheckout=1"
  
        $file = Get-File
    
    $body = @{
        attachments = @($file)
    } | ConvertTo-Json -Depth 5
    
     $testBody = [System.Text.Encoding]::UTF8.GetBytes($body)
$logBody = @{
    attachments = @(
        $file | Select-Object name
    )
} | ConvertTo-Json -Depth 5
       Write-MessageLog "[Line 121] Attachment add to Instance Param {COLUMN_Value:DET_Att5#DET_Att5} wfd_id $logBody"
     try{
       $response =  Invoke-RestMethod -Method Patch -Uri $newAtt -Body $testBody  -Headers $headers 
    
    Write-MessageLog "[Line 125] Response for wfd_id {COLUMN_Value:DET_Att5#DET_Att5} : $($response | ConvertTo-Json -Depth 5)"
    }catch{
          Write-MessageLog "[Line 126] Error : $_" -Level "ERROR"
        return $null
     }
    }
    add-Attachemts -baseUrl $baseUrl -clientId $CLIENT_ID -clientSecret $CLIENT_SECRET