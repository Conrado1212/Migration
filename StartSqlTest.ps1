$LogFile = "log.txt"
Clear-Content -Path $logFile -ErrorAction SilentlyContinue
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

#load env variables

#db Server name
$server = ""
#base Portal Url
$baseUrl = ""
#db name
$db = ""
#query to get data 
$query =""
function Get-Config{
    param(
        [string]$baseUrl,
        [string]$server,
        [string]$db,
        [string]$query
    )
  
     foreach ($p in $MyInvocation.MyCommand.Parameters.Keys) { 
         if (-not $PSBoundParameters.ContainsKey($p)) {
              Set-Variable -Name $p -Value (Read-Host "Add vealu for: $p")
             }
             } 
            
}
Get-Config


#function get data 
function Get-Data{
    param(
    [string]$server,
    [string]$db,
    [string]$query,
    [string]$LogPath = "",
    [int]$QueryTimeout = 30,
    [int]$ConnectionTimeout = 60
    )

    $logDir = Split-Path $LogPath 
    Write-MessageLog "[Line 63] $logDir"
    if (-not (Test-Path $logDir)) {
         New-Item -ItemType Directory -Path $logDir | Out-Null 
        }
   
    try {
        $result =    Invoke-Sqlcmd -ServerInstance $server -Database $db -Query $query -QueryTimeout $QueryTimeout -ConnectionTimeout $ConnectionTimeout
        Add-Content -Path $LogPath -Value "$(Get-Date) | SUCCESS | Query executed on $server/$db"
         return $result
    } catch {
        Add-Content -Path $LogPath -Value "$(Get-Date) | ERROR | $_"

        Write-MessageLog "Error  $_" -Level "ERROR"
       
    }

}






#Get-Data -server $server -db $db -Query $query 

#$dane | Format-Table -AutoSize

function Get-normalizeData{
   $dane =  Get-Data -server $server -db $db -Query $query 
 #  $dane | Format-Table -AutoSize
$resultNorm = @()
#decision
$dec = @{}
#map
$test = @()
$att = @()
foreach($item in $dane){
#Write-MessageLog "Line [86] $($item.nda_id)"
foreach($prop in $item.PSObject.Properties){
    if($prop.Name -like "nda_*"){
        if($dec.ContainsKey($prop.Name)){
            continue
        }
        $userInput = Read-Host "[Line 108] Do you want to add this $($prop.Name) to map (Y/Z?N)"
        if($userInput -eq "Y"){
            $dec[$prop.Name] = 'Y'
            $test += $prop.Name
        }elseif($userInput -eq "Z"){
            $dec[$prop.Name] = 'Z'
            $att +=$prop.Name
        }else{
            $dec[$prop.Name] = 'N'
        }
    }
}
}
foreach($item in $dane){
    $test2 =@{}
    foreach($prop in $test){
        if($item.PSObject.Properties[$prop]){
            $test2[$prop] = $item.prop
        }else{
            $test2[$prop] = $null
        }
    }
    $test2['attachments'] = @()
    $current = [PSCustomObject]$test2

        foreach($data in $att){

   if (![string]::IsNullOrWhiteSpace($item.$data) -and (Test-Path  $item.n$data)) {   
    try{
 $bytes = [System.IO.File]::ReadAllBytes($item.$data)
                    $base64 = [Convert]::ToBase64String($bytes)
                #nazwa pliku ze sciezki 
                 $name = Split-Path -Path  $item.$data -Leaf
                # Write-MessageLog "Nazwa pliku: $name ,base64:  $base64"
           $current.attachments == @{
               name = $name
               content = $base64
           }
    }catch{
        Write-MessageLog ("File {0}. Error: {1}" -f $item.$data, $_.Exception.Message) -Level "WARNING"
    }
                 
                  }
                }
                  $resultNorm +=$current
}
return $resultNorm
}

#Get-normalizeData

function Get-AccessToken{
    #parametr base url ktory podaje user
    #pobranei clientId oraz cleitnsecret jako zmiennej srodwoiskowej 
        param(
            [string]$baseUrl,
            [string]$clientId,
            [string]$clientSecret
        )
        if ([string]::IsNullOrWhiteSpace($clientId)) {

        $clientId = Read-Host "[Line 62] Write  CLIENT_ID"
        }else{
         Write-Host  "[Line 64] Variable clientId value retrieved"
        }

        if ([string]::IsNullOrWhiteSpace($clientSecret)) {
             $clientSecret = Read-Host "Write CLIENT_SECRET"
        }else{
          Write-Host   "[Line 70] Variable clientSecret value retrieved"
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
            Write-Error "[Line 88] Error authentication: $_"
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
        Write-Error "[Line 103] No token"
        return
    }

    $headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
    }
    $new = "$baseUrl/api/data/v6.0/db/1/elements?path="
    
   $dane =  Get-Data -server $server -db $db -Query $query
 

    foreach($item in $dane){
        
        $body =@{
            workflow = @{
                guid = ""
            }
            formType = @{
                guid = ""
            }
            formFields = @(
                @{
                     guid = ""   #1
                    svalue = "$($item.col2)"
                },
                @{
                     guid = "" #  2 
                    svalue = "$($item.col2)"
                },
                @{
                     guid = "" #  3 
                     svalue = "$($item.col3)"
                },
                @{
                    guid = "" #  5
                     svalue = "$($item.col5)"
                },
                @{
                     guid = "" #  6
                      svalue = "$($item.col8)"
                },
                @{
                     guid = "" #  7
                    svalue = "$($item.col9)"
                }
            )
            attachments = $item.attachments
            

           } | ConvertTo-Json -Depth 5
          # $allGroupsParam +=$body
          $testBody = [System.Text.Encoding]::UTF8.GetBytes($body)
          
          # Write-Host "[Line 121] New Instance Params $Body"
           Write-Host "[Line 163] New Instance Params $testBody"
            try{
          # $response =  Invoke-RestMethod -Method Post -Uri $new -Body $Body  -Headers $headers 
           $response =  Invoke-RestMethod -Method Post -Uri $new -Body $testBody  -Headers $headers 
            Write-Host "[Line 167] Instance created for id $($item.nda_id) $($response.id) " 
           # Write-Host "[Line 168] Response: $($response | ConvertTo-Json -Depth 5)"
        }catch{
        Write-Error "[Line 170] Error : $_"
        return $null
    }
    }  
        # $response | ConvertTo-Json -Depth 5
}
#$response = @(Start-New -baseUrl $baseUrl -clientId $CLIENT_ID -clientSecret $CLIENT_SECRET)

#Write-Host "[line 178] data: $response"