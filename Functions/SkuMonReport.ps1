Function SkuMonReport {
    [CmdletBinding()]
    param (
        #JSON settings file
        [Parameter(Mandatory, Position = 0)]
        [string]
        $ConfigFile,

        #skuCollection object passed from SkuMonGet
        [Parameter(Mandatory, Position = 1)]
        $skuCollection,

        #MS Graph Authentication Token Object passed from SkuMonGet
        [Parameter(Mandatory, Position = 2)]
        $Token
    )


    if (!(Test-Path $ConfigFile)) {
        Write-Warning "The specified configuration file cannot be found."
        EXIT
    }
    $resourceFolder = ((Split-Path -Path (Resolve-Path $PSScriptRoot).Path -Parent) + '\Resource')
    $css = Get-Content $resourceFolder\style.css -Raw
    $logo = [convert]::ToBase64String((Get-Content $resourceFolder\logo.png -Raw -Encoding byte))
    
    $timeZoneInfo = [System.TimeZoneInfo]::Local
    $tz = $timeZoneInfo.DisplayName.ToString().Split(" ")[0]

    $today = Get-Date -Format g
    
    #import config json
    Write-Verbose 'Importing configuration'
    $settings = (Get-Content $ConfigFile | ConvertFrom-Json)

    #create the final list for reporting.
    Write-Verbose 'Creating the list of Skus to be reported'
    $finalList = @()
    foreach ($item in ($settings.licenseToCheck | Where-Object { $_.Visible -eq $true })) {
        $sku = $skuCollection | Where-Object { $_.SkuId -eq $item.SkuID }
        if ($sku) {
            
            $skurow = New-Object psobject -Property ([ordered]@{
                    Name      = $item.SkuFriendlyName
                    Available = $sku.Available
                    Assigned  = $sku.Assigned
                    Total     = $sku.Total
                    Threshold = $item.Threshold
                    Status    = ( invoke-command {
                            if ($sku.Available -le $item.Threshold) {
                                return "Warning"
                            }
                            elseif ($sku.Available -gt $item.Threshold) {
                                return "Normal"
                            }
                        })
                })
            $finalList += $skurow
        }
        
    }    
    $finalList = $finalList | Sort-Object Threshold -Descending
    $reportCsv = ($settings.outputDirectory + '\SkuMonReport.csv')
    $finalList | Export-Csv -NoTypeInformation -Path $reportCsv

    #build HTML report
    Write-Verbose 'Start building HTML report'
    #$mailSubject = 'Microsoft 365 License Availability Report - ' + $today + ' ' + $tz
    $mailSubject = 'Microsoft 365 License Availability Report'
    $html = @()
    $html += '<html><head><title>' + $mailSubject + '</title>'
    $html += '<style type="text/css">'
    $html += $css
    $html += '</style></head>'
    $html += '<body>'
    #table headers
    $html += '<table id="tbl">'
    $html += '<tr><td class="head">' + $settings.organizationName + '<br>' + $today + ' ' + $tz +'</td></tr>'
    #$html += '<tr><td class="clean">' + $today + ' ' + $tz + '</td></tr>'
    $html += '<tr><td class="head"> </td></tr>'
    $html += '<tr><th class="section">Licenses</th></tr>'
    $html += '<tr><td class="head"> </td></tr>'
    $html += '<tr><td class="head"> </td></tr>'
    $html += '</table>'
    $html += '<table id="tbl">'
    $html += '<tr><td width="420px" colspan="2">Name</th><td width="170px">Available quantity</td><td width="5px"></td></tr>'
    
    foreach ($item in $finalList) {
        #$html += '<tr><td><img src="'+($resourceFolder+'\logo.png')+'" width="44" height="51"></img></td>'
        $html += '<tr><td><img src="'+($resourceFolder+'\logo.png')+'"></img></td>'
        $html += '<th>' + $item.Name + '</th>'
        $html += '<td>' + $item.Available + ' available<br>' + $item.Assigned + ' assigned of ' + $item.Total + ' total'
        if ($item.Threshold -ne 0) {
            if ($item.Status -eq 'Normal') {
                $html += '<td class="green" width="5px"></td></tr>'                
            }
            elseif ($item.Status -eq 'Warning') {
                $html += '<td class="red" width="5px"></td></tr>'                
            }
        }
        else {
            $html += '<td class="gray" width="5px"></td></tr>'
        }
    }
    $html += '<tr><td class="head" colspan="4"></td></tr>'
    $html += '</table>'

    $html += '<table id="legend">'
    $html += '<tr><td class="green" width="60px">Normal</td><td class="red" width="60px">Warning</td><td class="gray" width="60px">Ignored</td></tr>'
    $html += '</table>'

    $html += '<table id="settings">'
    $html += '<tr><td>Source:</td><td>' + $env:COMPUTERNAME + '</td></tr>'
    $html += '<tr><td>Setting:</td><td>' + (Resolve-Path $ConfigFile) + '</td></tr>'
    $html += '<tr><td colspan="2"><a href="https://github.com/junecastillote/LazyExchangeAdmin.SkuMon">Microsoft 365 Subscribed Sku Monitor</a></td></tr>'
    $html += '</table>'
    $html += '</body>'
    $html += '</html>'
    $reportHTML = ($settings.outputDirectory + '\SkuMonReport.html')
    $html = ($html -join "`n")
    $html | Out-File $reportHTML -Encoding utf8
    Write-Verbose ('HTML Report saved in ' + $reportHTML)

    if ($settings.mailSettings.sendEmail -eq $true) {
        Write-Verbose 'Sending email report'
        $mailSendURI = "https://graph.microsoft.com/v1.0/users/$($settings.mailSettings.From)/sendmail"

        $html = $html.Replace(($resourceFolder+'\logo.png'),"cid:logo")
        #construct To recipients
        $toAddressJSON = @()
        $settings.mailSettings.To | ForEach-Object {
            $toAddressJSON += @{EmailAddress = @{Address = ($_).Trim() } }
        }

        #construct CC recipients
        if ($settings.mailSettings.ccEnabled -eq $true) {
            $ccAddressJSON = @()
            $settings.mailSettings.cc | ForEach-Object {
                $ccAddressJSON += @{EmailAddress = @{Address = ($_).Trim() } }
            }
        }
        #construct BCC recipients
        if ($settings.mailSettings.bccEnabled -eq $true) {
            $bccAddressJSON = @()
            $settings.mailSettings.bcc | ForEach-Object {
                $bccAddressJSON += @{EmailAddress = @{Address = ($_).Trim() } }
            }
        }
   
        #build JSON mail payload
        $mailBody = @{
            message = @{
                subject                = $mailSubject
                body                   = @{
                    contentType = "HTML"
                    content     = $html
                }
                toRecipients           = @(
                    $ToAddressJSON
                )
                internetMessageHeaders = @(
                    @{
                        name  = "X-Mailer"
                        value = "skuMon by june.castillote@gmail.com"
                    }
                )
                attachments            = @(
                    @{
                        "@odata.type"  = "#microsoft.graph.fileAttachment"
                        "contentID"    = "logo"
                        "name"         = "logo"
                        "IsInline"     = $true
                        "contentType"  = "image/png"
                        "contentBytes" = $logo
                    }
                )
            }							
        }
        #add CC recipients
        if ($settings.mailSettings.ccEnabled -eq $true) {
            $mailBody.message += @{ccRecipients = $ccAddressJSON }
        }

        #add BCC recipients
        if ($settings.mailSettings.bccEnabled -eq $true) {
            $mailBody.message += @{bccRecipients = $bccAddressJSON }
        }
        
        $mailBody = $mailBody | ConvertTo-JSON -Depth 4
        try {
            Invoke-RestMethod -Method Post -Uri $mailSendURI -Body $mailbody -Headers $Token -ContentType application/json | out-null
            Write-Verbose 'Sent'
        }
        catch {
            Write-Error $_.Exception
            EXIT
        }        
    }
}