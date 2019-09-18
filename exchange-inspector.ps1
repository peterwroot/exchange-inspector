# EmailAudit.ps1
# A PowerShell script to find mailbox information on Microsoft Exchange servers. Also support Exchange Online (Office 365.)
function configureTerminal($width,$height,$window_title){

    # Function to customize the PowerShell host window. Takes vairables for the width,
    # height, and title text.

    $host.UI.RawUI.BackgroundColor = "White"
    $host.ui.RawUI.ForegroundColor = "Black"
    $windowSize = New-Object System.Management.Automation.Host.Size($width,$height)
    $host.ui.rawui.WindowSize=$windowSize   
    $host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size($width,$height)
    $host.ui.RawUI.WindowTitle = $window_title
}
function ChooseEmailServiceType {

    # This function presents the user with a menu to choose between the two types of supported email services:
    # Local Exchange and Office 365.

    Clear-Host
    $Title = "Microsoft Exchange / Office 365 Audit Tool"
    $Description = "Please choose the email server type:"
    $AuditTypeLocalExchange = New-Object System.Management.Automation.Host.ChoiceDescription "&Exchange","For On-Premesis Microsoft Exchange Servers"
    $AuditTypeOffice365 = New-Object System.Management.Automation.Host.ChoiceDescription "&Office 365","For Microsoft Office 365 (Exchange Online)"
    $AuditOptions =[System.Management.Automation.Host.ChoiceDescription[]]($AuditTypeLocalExchange,$AuditTypeOffice365)
    $result = $Host.UI.PromptForChoice($Title,$Description,$Auditoptions,1)
    return $result
}
Function TestCommand ($Command){

    # Tests to see if a command is valid. Used to see if the connection to Exchange
    # has already been made. 

    Try{
        Get-command $command -ErrorAction Stop
        Return $True
    }
    Catch [System.SystemException]{
        Return $False
    }
}
function ExportArrayToCSVFile($file_name, $array, $mailbox_type) {
    
    # Creates the output file using the provided name and the current datetime. 
    # Then writes the headers before writing the array containing the mailbox data.
    # The location of the file is then written to the terminal. 

    $exportDateTime = get-date -format "yyyyMMddHHmmss"
    $file_name = $exportDateTime + "-" + $file_name
    $FullFilePath = "C:\$file_name"
    New-Item $FullFilePath -Type File | Out-Null
    Add-Content $FullFilePath "Primary SMTP Address,Alias Addresses"
    foreach($line in $array){
        Add-Content $FullFilePath $line | out-null
    }
    if($mailbox_type.length -gt 14){
        write-host($mailbox_type + " audit: " + "`t" + $FullFilePath)
    }else{
        write-host($mailbox_type + " audit: " + "`t`t" + $FullFilePath)
    }   
}
function PressAnyKeyToContinue {

    # This function prints a 'Press any key to continue...' message to the terminal, and then continues the 
    # script when the user presses any key.

    write-host("")
    Write-Host -NoNewline 'Press any key to continue...'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Write-Host("`n")
}
function FormatEmailAddressString($EmailAddressList, $PrimarySMTPAddress) {

    # For each address, loop through and remove any address that use the SIP or SPO protocols (we only car about SMTP).
    # We then loop through the remaining addresses, and strip off the protocol headings at the beginning of the string. 
    # Then, we check if the address is the user's primary address, by comparing it to the $UserMailboxArray. It it is, we
    # add it to the $FormattedAddressString. Any remaining addresses get added to the $AliasList variable, along with a comma.
    # Next, we concatenate the $FormattedAddressString and the $AliasList to give us a single string, which starts with the
    # main mailbox address and then lists any aliases. Finally, we strip the trailing comma, and add the whole string to the 
    # $AliasAddressArray. BOOM. 

    $AliasList = ""
    $FormattedAddressString = ""

    foreach ($address in $EmailAddressList){
        if($address -notmatch "SIP:" -and $address -notmatch "SPO:" -and $address -notmatch "X500:"){
            $address = $address -replace "smtp:",""
            if($address -match $PrimarySMTPAddress){
                $FormattedAddressString = $address + ","
            }else{
                $AliasList += $address + ","
            }
        }
    }
    $FormattedAddressString += $AliasList
    $FormattedAddressString = $FormattedAddressString.TrimEnd(",")
    return $FormattedAddressString
}
function ConnectToOffice365 {

    # This function downloads and installs Microsoft's Exchange Online PowerShell Module. This
    # allows us to connect to Office 365, and supports 2FA. We download and execute the module,
    # which presents the user with an Office 365 login screen. Once the screen is done, we pass
    # control over to the AuditOffice365 function.

    Clear-Host
    $ExchangeOnlinePowerShellModuleURL = "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"

    Write-host("Step 1 of 4")
    write-host("------------------------------------------------------------------------")
    write-host("The Office 365 PowerShell Module will now be downloaded and installed.")
    Write-host("Please follow the installation wizard and accept any on-screen prompts.")

    PressAnyKeyToContinue
    start-process iexplore.exe $ExchangeOnlinePowerShellModuleURL
    clear-host

    # The Exchange Online PowerShell Module opens another PowerShell process when it
    # is finished installing. This loop runs until it finds another PowerShell process
    # with a different PID, and then kills it before continuing. 

    $pidkillcount = 0
    do {
        get-process -ProcessName PowerShell | ForEach-Object {
            if ($_.ID -ne $pid){
                stop-process -id $_.ID
                $pidkillcount += 1
            }
        }
    } until ($pidkillcount -ne 0)

    Write-host("Step 2 of 4")
    write-host("------------------------------------------------------------------------")
    Write-host("The Office 365 PowerShell Module is now installed.")
    write-host("You will now be presented with a login screen. Please log in as an")
    write-host("administrative Office 365 user.")

    PressAnyKeyToContinue
    
    # Locates the Exchange Online PowerShell Module in the user's AppData folder, then executes it.

    $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
    ."$MFAExchangeModule" 
    Clear-Host
    Connect-EXOPSSession    
    clear-host

    Write-host("Step 3 of 4")
    write-host("------------------------------------------------------------------------")
    write-host("You are now connected to Office 365. The mailbox audit will now begin.")

    PressAnyKeyToContinue
    clear-host
    AuditOffice365
}
function AuditOffice365 {

    # This function asks for infomation about each mailbox type we want to audit: User, Shared, and Distribution.
    # information about these mailboxes (their primary SMTP address, and then all of the mailbox's addresses) and
    # then passes that information over to the FormatEmailAddressString function to pretty up the output. The returned
    # strings are then passed to the ExportArrayToCSVFile function, and the total count of items in each array is
    # written to the terminal. Finally, the PSSession is ended, which disconnects the user from Office 365.

    Write-host("Step 4 of 4")
    write-host("------------------------------------------------------------------------")
    write-host("Auditing Mailboxes....")

    $UserMailboxAddressArray = @()
    $SharedMailboxArray = @()
    $DistributionGroupArray = @()

    get-mailbox -recipienttypedetails UserMailbox | ForEach-Object{
        $UserMailboxAddressArray += FormatEmailAddressString $_.emailaddresses $_.PrimarySMTPAddress
    }

    get-mailbox -recipienttypedetails SharedMailbox | foreach-object{
        $SharedMailboxArray += FormatEmailAddressString $_.emailaddresses $_.PrimarySMTPAddress
    }

    get-distributiongroup | ForEach-Object{
        $DistributionGroupArray += FormatEmailAddressString $_.emailaddresses $_.PrimarySMTPAddress
    }

    get-unifiedgroup | foreach-object{
        $DistributionGroupArray += FormatEmailAddressString $_.emailaddresses $_.PrimarySMTPAddress
    }
    
    write-host("Mailbox Audit Complete!" + "`n")

    write-host("Total User Mailboxes:" + "`t`t" + $UserMailboxAddressArray.count)
    Write-host("Total Shared Mailboxes:" + "`t`t" + $SharedMailboxArray.count)
    write-host("Total Distribution Groups:" + "`t" + $DistributionGroupArray.count + "`n")
    
    ExportArrayToCSVFile "UserMailboxes.csv" $UserMailboxAddressArray "User Mailbox"
    ExportArrayToCSVFile "SharedMailboxes.csv" $SharedMailboxArray "Shared Mailbox"
    ExportArrayToCSVFile "DistributionGroups.csv" $DistributionGroupArray "Distribution Group"

    Get-PSSession | Remove-PSSession
}
function ConnectToLocalExchange {

    # This function connects to the installation of Microsoft Exchange running on the machine that
    # the script is executed on. It then retireves the version of the Exchange Server and passes 
    # that information over to the AuditLocalExchange function so that it runs the correct commands
    # for the given Exchange Server version.

    Clear-Host
    
    Write-host("Step 1 of 2")
    write-host("------------------------------------------------------------------------")
    Write-host("This script will now connect to the installation of Exchange running on")
    Write-host("This computer.")

    PressAnyKeyToContinue
    Clear-Host
    
    if (-not(TestCommand "Get-Mailbox")){
        Invoke-Expression ". '$env:ExchangeInstallPath\Bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell"
    }
   
    Clear-Host
    
    $ExchangeMajorVersion = Get-ExchangeServer | Select-Object -ExpandProperty AdminDisplayVersion | Select-Object -ExpandProperty Major
    $ExchangeMinorVersion = Get-ExchangeServer | Select-Object -ExpandProperty AdminDisplayVersion | Select-Object -ExpandProperty Minor
    $ExchangeVersion = "$ExchangeMajorVersion.$ExchangeMinorVersion"

    switch($ExchangeVersion){
        "15.2" {$ExchangeVersionYear = "2019"}    # Exchange Server Version: 2019
        "15.1" {$ExchangeVersionYear = "2016"}    # Exchange Server Version: 2016
        "15.0" {$ExchangeVersionYear = "2013"}    # Exchange Server Version: 2013
        "14.3" {$ExchangeVersionYear = "2010"}    # Exchange Server Version: 2010
        "14.2" {$ExchangeVersionYear = "2010"}    # Exchange Server Version: 2010
        "14.1" {$ExchangeVersionYear = "2010"}    # Exchange Server Version: 2010
        "14.0" {$ExchangeVersionYear = "2010"}    # Exchange Server Version: 2010
        "8.3"  {$ExchangeVersionYear = "2007"}    # Exchange Server Version: 2007
        "8.2"  {$ExchangeVersionYear = "2007"}    # Exchange Server Version: 2007
        "8.1"  {$ExchangeVersionYear = "2007"}    # Exchange Server Version: 2007
        "8.0"  {$ExchangeVersionYear = "2007"}    # Exchange Server Version: 2007
    }

    Clear-Host
    configureTerminal 72 10 "Email Service Audit"

    Write-host("Step 2 of 2")
    write-host("------------------------------------------------------------------------")
    Write-host("You are now connected to Exchange on $env:COMPUTERNAME.$env:userdnsdomain")
    Write-host("Detected Exchange Server Version: " + "`t" + $ExchangeVersionYear)
    write-host("Detected Exchange Server Build: " + "`t" + $ExchangeVersion)  
    write-host("The mailbox audit will now begin. ")

    PressAnyKeyToContinue

    AuditLocalExchange $ExchangeVersionYear
}
function AuditLocalExchange ($version){
    
    # When I get around to writing this function, it will audit the specified exchange server.
    clear-host
    #start-process iexplore.exe "https://i.imgflip.com/2q5hiq.jpg"

    $UserMailboxAddressArray = @()
    $SharedMailboxArray = @()
    $DistributionGroupArray = @()

    get-mailbox -recipienttypedetails UserMailbox | ForEach-Object{
        $UserMailboxAddressArray += FormatEmailAddressString $_.emailaddresses $_.PrimarySMTPAddress
    }

    get-mailbox -recipienttypedetails SharedMailbox | foreach-object{
        $SharedMailboxArray += FormatEmailAddressString $_.emailaddresses $_.PrimarySMTPAddress
    }

    get-distributiongroup | ForEach-Object{
        $DistributionGroupArray += FormatEmailAddressString $_.emailaddresses $_.PrimarySMTPAddress
    }

    write-host("Mailbox Audit Complete!" + "`n")

    write-host("Total User Mailboxes:" + "`t`t" + $UserMailboxAddressArray.count)
    Write-host("Total Shared Mailboxes:" + "`t`t" + $SharedMailboxArray.count)
    write-host("Total Distribution Groups:" + "`t" + $DistributionGroupArray.count + "`n")
    
    ExportArrayToCSVFile "UserMailboxes.csv" $UserMailboxAddressArray "User Mailbox"
    ExportArrayToCSVFile "SharedMailboxes.csv" $SharedMailboxArray "Shared Mailbox"
    ExportArrayToCSVFile "DistributionGroups.csv" $DistributionGroupArray "Distribution Group"
}

# Prompts the user to choose which email service type to use.
configureTerminal 72 10 "Email Service Audit"

switch (ChooseEmailServiceType) {
    0 {ConnectToLocalExchange}
    1 {ConnectToOffice365}
    Default {}
}