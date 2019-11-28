Function SkuMonToken {
    param (
        [parameter(Mandatory, Position = 0)]
        $ClientID,
        [parameter(Mandatory, Position = 1)]
        $ClientSecret,
        [parameter(Mandatory, Position = 2)]
        $TenantID
    )
    $body = @{grant_type = "client_credentials"; scope = "https://graph.microsoft.com/.default"; client_id = $ClientID; client_secret = $ClientSecret }
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token -Body $body
    $token = @{'Authorization' = "$($oauth.token_type) $($oauth.access_token)" }
    return $token
}

Function SkuMonLogStop
{
	$txnLog=""
	Do {
		try {
			Stop-Transcript | Out-Null
		}
		catch [System.InvalidOperationException]{
			$txnLog="stopped"
		}
    } While ($txnLog -ne "stopped")
}

#Function to Start transcribing
Function SkuMonLogStart
{
    param
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$logDirectory
    )
	SkuMonLogStop
    Start-Transcript $logDirectory -Append | Out-Null
}