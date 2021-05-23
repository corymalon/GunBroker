##################################################################
#                                                                #
# This code is provided for use by the community free of charge. #
# The code is offered with no warranties, express or implied.    #
#                                                                #
##################################################################


clear

# Checks to see if credential files are in tact. If you've changed your password for GunBroker, delete the PWD.txt file and you will be prompted to re-enter it
$UIDTest = Test-Path -Path .\UID.txt
$PWDTest = Test-Path -Path .\PWD.txt

if ($UIDTest -eq $false){
    $UID = Read-Host "Enter Your GunBroker Username"
    Add-Content .\UID.txt -Value $UID
}

if ($PWDTest -eq $FALSE){
    $PWD = Read-Host "Enter Your GunBroker Password" -AsSecureString | ConvertFrom-SecureString | Out-File ".\PWD.txt"
}

# Pulls credentials for authentication to GunBroker
$User = Get-Content ".\UID.txt"
$pw = Get-Content ".\PWD.txt"

# Decrypts password in order to pass to the GunBroker website
$MyCredential=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, ($pw | ConvertTo-SecureString)
$pass = $MyCredential.getnetworkcredential().Password

# Pepares credentials for login to GunBroker
$postParams = @{Username=$User;Password=$pass}

# Executes the login to GunBroker and establishes a session on the website
$gblogin = Invoke-WebRequest -uri "https://www.gunbroker.com/user/login" -Method POST -Body $postParams -SessionVariable 'Session'

# Imports the lookup list
$ImportList = Import-Csv ".\Lookup.csv"

# Creates the output file
$DateTime = Get-Date -Format MMddyy-HHmmss
$OutFile = "Output$DateTime.csv"
Add-Content -path .\$OutFile -Value "Make,Model,Caliber,Price"

# Returns average sale values based on import list
ForEach ($Import in $ImportList){
    # Converts human readable list into something machines can read
    $urlString = $Import.make+" "+$Import.model+" "+$Import.caliber
    $encode = [uri]::EscapeDataString($urlString)

    # Executes a search of completed listings
    $gb = Invoke-WebRequest -uri "https://www.gunbroker.com/All/search/completed?Keywords=$encode&PageSize=96&Sort=1&View=1&Condition=4&Timeframe=1" -WebSession $Session

    # Scrapes the return of completed listings and converts the string output into a decimal value for so the computer can do math
    $PriceList = $gb.ParsedHtml.body.getElementsByClassName("current") | Select -ExpandProperty InnerHtml 
    $PriceList = $PriceList -replace ',', ''
    $PriceList = $PriceList -replace '\$',''

    # Logic to check whether or not results for the GunBroker search exist or not
    if ($PriceList -ne ""){
        # Calculates the average value of the item in the list
        $avg = $PriceList | Measure-Object -average | select -ExpandProperty average
        $price = [math]::Round($avg,2)
        $Output = $Import.make+","+$Import.model+","+$Import.caliber+","+$price
        Add-Content -path .\$OutFile -Value $Output
    }
    else {
        # Write this out if no results exist
        $Output = $Import.make+","+$Import.model+","+$Import.caliber+","+"NO RESULTS"
        Add-Content -path .\$OutFile -Value $Output
    }    
}