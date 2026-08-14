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
    
    $testFiles = @("{COLUMN_Value:DET_Att6#DET_Att6}", "{COLUMN_Value:DET_Att6#DET_Att7}");

    foreach ($testFile in $testFiles) {
        if ([string]::IsNullOrWhiteSpace($testFile)) {
            continue
        }

     $bytes = [System.IO.File]::ReadAllBytes($testFile)
      $name = Split-Path -Path  $testFile -Leaf
       $att +=[PSCustomObject]@{
                            Name = $name
                            Bytes = [Convert]::ToBase64String($bytes)
                        }
                    }
                        return $att
}
#$file = Get-File
#Write-Output $file
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
$newAtt = "$baseUrl/api/data/v6.0/db/1/elements/{COLUMN_Value:DET_Att5#DET_Att5}/attachments?forceCheckout=1"
    $file = Get-File

$body = @{
    name = $file.Name
    content = [Convert]::ToBase64String($file.Bytes)
} | ConvertTo-Json -Depth 5

 $testBody = [System.Text.Encoding]::UTF8.GetBytes($body)
  # Write-MessageLog "[Line 121] Attachment add to Instance Param $testBody"
 try{
   $response =  Invoke-RestMethod -Method Post -Uri $newAtt -Body $testBody  -Headers $headers 

Write-MessageLog "[Line 125] Response for wfd_id {COLUMN_Value:DET_Att5#DET_Att5} : $($response | ConvertTo-Json -Depth 5)"
}catch{
      Write-MessageLog "[Line 126] Error : $_" -Level "ERROR"
    return $null
 }
}
 Start-New -baseUrl $baseUrl -clientId $CLIENT_ID -clientSecret $CLIENT_SECRET


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
   # $newAtt = "$baseUrl/api/data/v6.0/db/1/elements/{COLUMN_Value:DET_Att5#DET_Att5}/attachments?forceCheckout=1"
    $newAtt = "$baseUrl/api/data/v6.0/db/1/elements/{COLUMN_Value:DET_Att5#DET_Att5}?forceCheckout=1"
  
        $file = Get-File
    
    $body = @{
        attachments = $file
    } | ConvertTo-Json -Depth 5
    
     $testBody = [System.Text.Encoding]::UTF8.GetBytes($body)
      # Write-MessageLog "[Line 121] Attachment add to Instance Param $testBody"
     try{
       $response =  Invoke-RestMethod -Method Post -Uri $newAtt -Body $testBody  -Headers $headers 
    
    Write-MessageLog "[Line 125] Response for wfd_id {COLUMN_Value:DET_Att5#DET_Att5} : $($response | ConvertTo-Json -Depth 5)"
    }catch{
          Write-MessageLog "[Line 126] Error : $_" -Level "ERROR"
        return $null
     }
    }
    add-Attachemts -baseUrl $baseUrl -clientId $CLIENT_ID -clientSecret $CLIENT_SECRET