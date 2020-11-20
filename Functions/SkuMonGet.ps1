
Function SkuMonGet {

    [cmdletbinding(DefaultParameterSetName = 'byToken')]
    param(
        [parameter(ParameterSetName = 'byToken', Mandatory)]
        $Token,
        [parameter(ParameterSetName = 'byID', Mandatory)]
        [string]$ClientID,
        [parameter(ParameterSetName = 'byID', Mandatory)]
        [string]$ClientSecret,
        [parameter(ParameterSetName = 'byID', Mandatory)]
        [string]$TenantID,
        [parameter(ParameterSetName = 'byID')]
        [string]$NewSetup,
        [parameter(ParameterSetName = 'byConfig', Mandatory)]
        [string]$ConfigFile
    )

    SkuMonLogStart $env:windir\temp\SkuMonLog.log
    #Get Token (byID)
    if ($PSCmdlet.ParameterSetName -eq 'byID') {
        try {
            Write-Verbose "Get MS Graph API Authorization Token"
            $token = SkuMonToken $ClientID $ClientSecret $TenantID
        }
        catch {
            Write-Host $_.Exception -ForegroundColor Red
            Return $null
        }
    }

    #Get Token (byConfig)
    if ($PSCmdlet.ParameterSetName -eq 'byConfig') {
        $settings = (Get-Content $ConfigFile -Raw | ConvertFrom-Json)
        $ClientID = $settings.appSettings.ClientID
        $ClientSecret = $settings.appSettings.ClientSecret
        $TenantID = $settings.appSettings.TenantID
        try {
            Write-Verbose "Get MS Graph API Authorization Token"
            $token = SkuMonToken $ClientID $ClientSecret $TenantID
        }
        catch {
            Write-Host $_.Exception -ForegroundColor Red
            SkuMonLogStop
            Return $null
        }
    }

    $skuTable = Import-Csv ((Split-Path -Path (Resolve-Path $PSScriptRoot).Path -Parent) + '\Resource\skuTable.csv')

    #Get List of Skus
    #Reference - https://docs.microsoft.com/en-us/graph/api/subscribedsku-list?view=graph-rest-1.0&tabs=http
    try {
        Write-Verbose "Get list of subscribed licenses"
        $uri = 'https://graph.microsoft.com/v1.0/subscribedSkus'
        $subscribedSkus = Invoke-RestMethod -Method GET -Uri $uri -Headers $Token
        $organization = invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $Token
        $organizationName = $organization.value.displayname
    }
    catch {
        Write-Host $_.Exception -ForegroundColor Red
        SkuMonLogStop
        Return $null
    }

    #process the collection
    Write-Verbose 'Start creating the Subscribed Sku List'
    $skuCollection = @()
    foreach ($sku in ($subscribedSkus.Value | Where-Object { $_.appliesTo -eq 'User' })) {
        #available and excess
        $AvailableUnits = (($sku.prepaidUnits.Enabled + $sku.prepaidUnits.Warning) - $sku.ConsumedUnits)
        $ExcessUnits = 0
        if ($AvailableUnits -lt 0) {
            $ExcessUnits = [Math]::Abs($AvailableUnits)
            $AvailableUnits = 0
        }
        #build properties
        $splat = [ordered]@{
            SkuID           = $sku.SkuID
            SkuPartNumber   = $sku.SkuPartNumber
            SkuFriendlyName = (($skuTable | Where-Object { $_.SkuID -eq $sku.SkuID }).SkuFriendlyName)
            Assigned        = $sku.ConsumedUnits
            Total           = $sku.prepaidUnits.Enabled
            Suspended       = $sku.prepaidUnits.Suspended
            Warning         = $sku.prepaidUnits.Warning
            Available       = $AvailableUnits
            Invalid         = $ExcessUnits
            Status          = $sku.CapabilityStatus
        }
        $tempObj = New-Object psobject -Property $splat
        $skuCollection += $tempObj
    }

    #$base64_Style = "I3RibCANCnsNCiAgICBmb250LWZhbWlseToiU2Vnb2UgVUkiOw0KICAgIHdpZHRoOjcwJTsNCiAgICBib3JkZXItY29sbGFwc2U6Y29sbGFwc2U7DQoJbWFyZ2luLWxlZnQ6YXV0bzsNCgltYXJnaW4tcmlnaHQ6YXV0bzsNCn0gDQojdGJsIHRkLCAjdGJsIHRoDQp7IA0KICAgIGZvbnQtc2l6ZToxNHB4Ow0KICAgIGJvcmRlcjoxcHggbm9uZSAjREREOw0KICAgIHBhZGRpbmctdG9wOjVweDsNCiAgICBwYWRkaW5nLWJvdHRvbTo1cHg7DQogICAgcGFkZGluZy1sZWZ0OjEwcHg7DQp9IA0KI3RibCB0aA0Kew0KICAgIGZvbnQtc2l6ZToxNHB4Ow0KICAgIGZvbnQtd2VpZ2h0OiBib2xkOw0KICAgIGJhY2tncm91bmQtY29sb3I6I2ZmZjsNCiAgICBjb2xvcjojMDAwOyB0ZXh0LWFsaWduOmxlZnQ7DQogICAgdmVydGljYWwtYWxpZ246dG9wOw0KfSANCiN0YmwgdGguc2VjdGlvbg0Kew0KICAgIGZvbnQtZmFtaWx5OiJTZWdvZSBVSSBMaWdodCI7DQogICAgZm9udC1zaXplOjI0cHg7DQogICAgdGV4dC1hbGlnbjpsZWZ0Ow0KICAgIHBhZGRpbmctdG9wOjEwcHg7DQogICAgcGFkZGluZy1ib3R0b206MjBweDsNCiAgICBwYWRkaW5nLWxlZnQ6MTBweDsNCiAgICBiYWNrZ3JvdW5kLWNvbG9yOiNmZmY7DQogICAgY29sb3I6IzAwMDsNCiAgICB2ZXJ0aWNhbC1hbGlnbjpjZW50ZXI7DQp9DQojdGJsIHRkIA0KeyANCiAgICBmb250LXNpemU6MTRweDsNCiAgICB0ZXh0LWFsaWduOmxlZnQ7DQogICAgdmVydGljYWwtYWxpZ246dG9wOw0KfSANCiN0YmwgdGQuYmFkDQp7DQogICAgZm9udC1zaXplOjE2cHg7DQogICAgZm9udC13ZWlnaHQ6IGJvbGQ7DQogICAgY29sb3I6I2YwNDk1MzsNCiAgICB2ZXJ0aWNhbC1hbGlnbjp0b3A7DQp9IA0KI3RibCB0ZC5nb29kDQp7DQogICAgZm9udC1zaXplOjE2cHg7DQogICAgZm9udC13ZWlnaHQ6IGJvbGQ7DQogICAgY29sb3I6IzAxYTk4MjsNCiAgICB2ZXJ0aWNhbC1hbGlnbjp0b3A7DQp9"
    #$css_string = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64_Style))

    if ($PSCmdlet.ParameterSetName -ne 'byConfig' -and !$NewSetup) {
        SkuMonLogStop
        return $skuCollection
    }

    if ($PSCmdlet.ParameterSetName -eq 'byID' -and $NewSetup) {

        Write-Verbose 'Setting values for the new setup file'
        $licenseToCheck = @()
        foreach ($item in $skuCollection) {
            $props = [ordered]@{
                Visible         = $true
                SkuID           = $item.SkuID
                SkuPartNumber   = $item.SkuPartNumber
                SkuFriendlyName = $item.SkuFriendlyName
                Threshold       = 0
            }
            $selection = New-Object psobject -property $props
            $licenseToCheck += $selection
        }

        if ($licenseToCheck.Count -eq 1) {
            $licenseToCheck += @{
                Visible = $false
                SkuID = ((New-Guid).Guid)
                SkuPartNumber = 'DUMMY'
                SkuFriendlyName = 'DUMMY'
                Threshold = 0
            }
        }

        $licenseToCheck = $licenseToCheck | Sort-Object License

        $props = @{
            organizationName = $organizationName
            appSettings      = @{
                clientID     = $ClientID
                clientSecret = $ClientSecret
                tenantID     = $TenantID
            }
            mailSettings     = @{
                sendEmail  = $false
                ccEnabled  = $false
                bccEnabled = $false
                From       = ""
                To         = @()
                Cc         = @()
                Bcc        = @()
            }
            licenseToCheck   = $licenseToCheck
            outputDirectory  = ($NewSetup).Split("\")[-1].Split(".")[0] + '-output'
        }
        $settings = new-object psobject -property $props

        if (!(Test-Path (Split-Path -Parent -Path $NewSetup))) {
            New-Item -ItemType Directory -Path (Split-Path -Parent -Path $NewSetup) | Out-Null
        }
        $settings.outputDirectory = (Resolve-Path (Split-Path -Parent -Path $NewSetup)).Path

        $settings | ConvertTo-Json | Out-File $NewSetup
        Write-Verbose 'New setup file created'
        Write-Verbose 'Opening the setup interface'
        $null = SkuMonSetup (Resolve-Path $NewSetup)
        Write-Verbose 'Done'
        SkuMonLogStop
    }

    if ($PSCmdlet.ParameterSetName -eq 'byConfig') {
        Write-Verbose 'Creating report'
        SkuMonReport $ConfigFile $skuCollection $Token
    }
    SkuMonLogStop
}
