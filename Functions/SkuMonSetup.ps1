
Add-Type -AssemblyName PresentationFramework
Function SkuMonSetup {
    param (
        #JSON Settings file.
        [Parameter()]
        [string]
        $ConfigFile
    )

    #create window
    $inputXML = Get-Content ((Split-Path -Path (Resolve-Path $PSScriptRoot).Path -Parent)+'\Resource\setup.xaml') -Raw
    #$inputXML = $inputXML -replace 'Binding="{x:Null}" ClipboardContentBinding="{x:Null}"', ''
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'

    [xml]$XAML = $inputXML
    #Read XAML

    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    try {
        $window = [Windows.Markup.XamlReader]::Load( $reader )
    }
    catch {
        Write-Warning $_.Exception
        throw
    }
    #Create variables based on form control names.
    #Variable will be named as 'var_<control name>'
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object { #"trying item $($_.Name)";
        try {
            Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
        }
        catch {
            throw
        }
    }
    #Get-Variable var_*

    Function OpenSettings {
        param (
            # Parameter help description
            [Parameter(Position = 0)]
            [string]
            $ConfigFile
        )
        if ($ConfigFile) {
            $global:settingsFile = ($ConfigFile)
            $window.Title = $windowTitle + ' - ' + ($global:settingsFile)
            $global:settings = Get-Content ($ConfigFile) | ConvertFrom-Json
            return $global:settings
        }
        else {
            $dialog = New-Object -TypeName 'Microsoft.Win32.OpenFileDialog'
            $dialog.Title = 'Open...'
            $dialog.Filter = 'JSON Settings|*.json'
            if ($dialog.ShowDialog() -eq $true) {
                $global:settingsFile = ($dialog.FileName)
                $window.Title = $windowTitle + ' - ' + ($global:settingsFile)
                $global:settings = Get-Content ($global:settingsFile) | ConvertFrom-Json
                return $global:settings
            }
            else {
                return $global:settings
            }
        }
    }

    Function GetLastSettingsSaveTime {
        Return (Get-Item $global:settingsFile).LastWriteTime
    }
    Function SaveAsSettings {
        $dialog = New-Object -TypeName 'Microsoft.Win32.SaveFileDialog'
        $dialog.Title = 'Save As...'
        $dialog.Filter = 'JSON Settings|*.json'
        if ($dialog.ShowDialog() -eq $true) {
            $global:settingsFile = ($dialog.FileName)
            SaveSettings $global:settingsFile
        }
    }
    #start SaveSettings
    Function SaveSettings {
        param (
            [Parameter(Position = 0, Mandatory)]
            [string]$filename
        )
        #validate fields
        $validationPassed = $true

        if (!$var_txtClientID.Text) {
            $validationPassed = $false
            Write-Warning "Missing ClientID"
        }
        if (!$var_txtClientSecret.Text) {
            $validationPassed = $false
            Write-Warning "Missing ClientSecret"
        }
        if (!$var_txtTenantID.Text) {
            $validationPassed = $false
            Write-Warning "Missing TenantID"
        }
        if (!$var_txtOrgName.Text) {
            $validationPassed = $false
            Write-Warning "Missing Organization Name"
        }

        if ($var_chkSendEmail.IsChecked -eq $true) {
            if (!$var_txtTo.Text) {
                $validationPassed = $false
                Write-Warning "Missing TO address"
            }
            else {
                foreach ($addr in ($var_txtTo.Text -split ",")) {
                    try {
                        $null = [mailaddress]$addr
                    }
                    catch {
                        $validationPassed = $false
                        Write-Warning "To address ($($addr)) invalid format"
                    }
                }
            }

            if (!$var_txtFrom.Text) {
                $validationPassed = $false
                Write-Warning "Missing FROM address"
            }
            else {
                try {
                    $null = [mailaddress]$var_txtFrom.Text
                }
                catch {
                    $validationPassed = $false
                    Write-Warning "From address ($($var_txtFrom.Text))  invalid format"
                }
            }

            if ($var_chkCC.IsChecked -eq $true -and !$var_txtCC.Text) {
                $validationPassed = $false
                Write-Warning "Missing CC address"
            }
            elseif ($var_txtCC.Text) {
                foreach ($addr in ($var_txtCC.Text -split ",")) {
                    try {
                        $null = [mailaddress]$addr
                    }
                    catch {
                        $validationPassed = $false
                        Write-Warning "CC address ($($addr)) invalid format"
                    }
                }
            }

            if ($var_chkBcc.IsChecked -eq $true -and !$var_txtBcc.Text) {
                $validationPassed = $false
                Write-Warning "Missing Bcc address"
            }
            elseif ($var_txtBcc.Text) {
                foreach ($addr in ($var_txtBcc.Text -split ",")) {
                    try {
                        $null = [mailaddress]$addr
                    }
                    catch {
                        $validationPassed = $false
                        Write-Warning "Bcc address ($($addr)) invalid format"
                    }
                }
            }
        }
        #========
        #store contents of dataGrid

        if ($validationPassed -eq $true) {
            $global:settings.licenseToCheck = $var_dGrid1.Items
            $global:settings.organizationName = $var_txtOrgName.Text
            $global:settings.appSettings.clientID = $var_txtClientID.Text
            $global:settings.appSettings.clientSecret = $var_txtclientSecret.Text
            $global:settings.appSettings.tenantID = $var_txtTenantID.Text
            $global:settings.mailSettings.sendEmail = $var_chkSendEmail.IsChecked
            $global:settings.mailSettings.ccEnabled = $var_chkCc.IsChecked
            $global:settings.mailSettings.bccEnabled = $var_chkBcc.IsChecked
            $global:settings.mailSettings.From = $var_txtFrom.Text
            $global:settings.mailSettings.To = ($var_txtTo.Text).Split(",")
            $global:settings.mailSettings.Cc = ($var_txtCc.Text).Split(",")
            $global:settings.mailSettings.Bcc = ($var_txtBcc.Text).Split(",")
            $global:settings.outputDirectory = (Split-Path -Parent -Path ($global:settingsFile)) + '\' + ((Split-Path $global:settingsFile -Leaf).Split(".")[0] + '-output')

            if (!(Test-Path ($global:settings.outputDirectory))) {
                New-Item -ItemType Directory -Path ($global:settings.outputDirectory)
            }
            #Save settings
            $global:settings | convertto-json | out-file $filename
            $var_lblStatus.Content = 'Saved. (' + (Get-Date -Format G) + ')'
            $window.Title = $windowTitle + ' - ' + $global:settingsFile
        }
        else {
            $var_lblStatus.Content = "Missing or invalid values detected. Please review your entries before saving again."
        }
    }
    #.....................................

    $var_lblStatus.Content = "Open a settings file."
    $windowTitle = "SkuMon - Setup"
    if ($ConfigFile) {
        if (Test-Path $ConfigFile) {
            $global:settingsFile = (Resolve-Path $ConfigFile).Path
            $global:settings = OpenSettings $ConfigFile
            if ($global:settings) {
                $var_txtOrgName.Text = $global:settings.organizationName
                $var_txtClientID.Text = $global:settings.appSettings.clientID
                $var_txtClientSecret.Text = $global:settings.appSettings.clientSecret
                $var_txtTenantID.Text = $global:settings.appSettings.tenantID
                $var_txtTo.Text = ($global:settings.mailSettings.To -Join ",")
                $var_txtFrom.Text = $global:settings.mailSettings.From
                $var_txtCc.Text = ($global:settings.mailSettings.Cc -Join ",")
                $var_txtBcc.Text = ($global:settings.mailSettings.Bcc -Join ",")
                $var_chkSendEmail.IsChecked = $global:settings.mailSettings.sendEmail
                $var_chkCc.IsChecked = $global:settings.mailSettings.ccEnabled
                $var_chkBcc.IsChecked = $global:settings.mailSettings.bccEnabled
                $var_dGrid1.ItemsSource = $global:settings.licenseToCheck | Sort-Object SkuFriendlyName
                $var_btnSave.IsEnabled = $true
                $var_grpMonitor.IsEnabled = $true
                $var_grpEmail.IsEnabled = $true
                $var_grpApp.IsEnabled = $true
                $var_txtClientID.Focus()
                $var_lblStatus.Content = 'Opened setting file - ' + $global:settingsFile
            }
        }
        else
        {
            Write-Host ($ConfigFile + ' cannot be found')
            Return $null
        }
    }
    else {
        $global:settingsFile = ""
    }

    #Add shortcut keys
    $commonKeyEvents = {
        [System.Windows.Input.KeyEventArgs] $e = $args[1]

        if ($e.Key -eq "S" -and $e.KeyboardDevice.Modifiers -eq "Ctrl") {
            SaveSettings $global:settingsFile
        }
    }
    #.....................................

    $window.Add_PreViewKeyDown($commonKeyEvents)

    #add control functions
    #SAVE button click
    $var_btnSave.Add_Click( {
            SaveSettings $global:settingsFile
        })

    #OPEN button click
    $var_btnOpen.Add_Click( {

            $global:settings = OpenSettings
            if ($global:settings) {
                $var_txtOrgName.Text = $global:settings.organizationName
                $var_txtClientID.Text = $global:settings.appSettings.clientID
                $var_txtClientSecret.Text = $global:settings.appSettings.clientSecret
                $var_txtTenantID.Text = $global:settings.appSettings.tenantID
                $var_txtTo.Text = ($global:settings.mailSettings.To -Join ",")
                $var_txtFrom.Text = $global:settings.mailSettings.From
                $var_txtCc.Text = ($global:settings.mailSettings.Cc -Join ",")
                $var_txtBcc.Text = ($global:settings.mailSettings.Bcc -Join ",")
                $var_chkSendEmail.IsChecked = $global:settings.mailSettings.sendEmail
                $var_chkCc.IsChecked = $global:settings.mailSettings.ccEnabled
                $var_chkBcc.IsChecked = $global:settings.mailSettings.bccEnabled
                $var_dGrid1.ItemsSource = $global:settings.licenseToCheck | Sort-Object SkuFriendlyName
                $var_btnSave.IsEnabled = $true
                $var_grpMonitor.IsEnabled = $true
                $var_grpEmail.IsEnabled = $true
                $var_grpApp.IsEnabled = $true
                $var_txtClientID.Focus()
                $var_lblStatus.Content = 'Opened setting file - ' + $global:settingsFile
            }
        })
    $Null = $window.ShowDialog()
}