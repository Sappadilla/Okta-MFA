<#
.SYNOPSIS
    This is the main script for administrating MFA in Okta.
.DESCRIPTION
    It has functions for:
        -Creating Google MFA Activation codes
        -Adding SMS MFA
        -Adding Voice Call MFA
        -Deleting Factors
        -Generate Google MFA codes from a CSV
    Input options:
        -Email (can get the email from NTID, or First Name and Last Name)
        -Phone Number (for voice call or SMS MFA)
        -CSV (will need to give the column names for email or NTID, and MFA.
            If column names for email and MFA are not given, the script will create them, but it must have
            some input for user values.)
.NOTES
    Sources : https://github.com/gabrielsroka/OktaAPI.psm1
              https://github.com/HumanEquivalentUnit/PowerShell-Misc/blob/master/GoogleAuthenticator.psm1
    Author  : Justin Sappington
    Created : October 4th, 2019
    Last Modified : November 6th, 2019
    Dependencies  : Powershell ActiveDirectory Module
                  : Namespace System
    Full Documentation: https://confluence.nike.com/display/IAM/Okta+MFA+Setup
    
    $PSVersionTable for where it was written:
        Name                           Value                                                                                          
        ----                           -----                                                                                          
        PSVersion                      5.1.15063.1805                                                                                 
        PSEdition                      Desktop                                                                                        
        PSCompatibleVersions           {1.0, 2.0, 3.0, 4.0...}                                                                        
        BuildVersion                   10.0.15063.1805                                                                                
        CLRVersion                     4.0.30319.42000                                                                                
        WSManStackVersion              3.0                                                                                            
        PSRemotingProtocolVersion      2.3                                                                                            
        SerializationVersion           1.1.0.1  
    
    Change Notes:
        Date: 10-25-2019 - Actor: Justin Sappington
        Description:
                    -Added "-NTID" flag to Add-Factor that will call Get-Email instead of being required
                    to be called at the command line input. Going from "Add-Factor -Google -Email (Get-Email -NTID $user_ntid)"
                    to "Add-Factor -Google -NTID $user_NTID".
                    -Added [Parameter(mandatory=$false)]$NTID="".  Changed $Email parameter to mandatory=$false.
                    -Added error handling for $NTID and $email being false
                    -Updated output text to be "factor" instead of "factors" when only one factor is setup by Add-Factor.
                    -Updated output text to be "factor" instead of "factors" when only one factor is removed by Remove-Factor.
                    -Added No_Connect flag to List-Factors
                    -Added -NTID flag for Remove-Factor function

        Date: 10-30-2019 - Actor: Justin Sappington
        Description:
                    -Added input filtering to Get-Email to remove white spaces around NTIDs.  This will filter for all
                    functions that call Get-Email instead of being performed in each individual function
                    -Added cleaner error handling for invalid phone numbers based on error code in Add-Factor
                    -Added error handling for 404 Not Found in List-Factors
        Date: 11-01-2019 - Actor: Justin Sappington
        Description:
                    -Changed Token storage from hard coded inside the script to being pulled from a config file in the
                    same directory as this script and the friendly setup script.
                    -Did the same for the base URL. Pulls from config file instead of being hard coded.
        Date: 11-06-2019 - Actor: Justin Sappington
        Description:
                    -Modified Add-Factor-CSV function.  Added handling for finding existing MFA code columns.  Searches
                    for any column with "MFA" in the name, then uses that column as the output column for the
                    MFA activation codes.

.EXAMPLE
    Get-Email -NTID "jsappi"
    Get-Email -First "Justin" -Last "Sappington"

    List-Factors -email Justin.Sappington@nike.com
    List-Factors -email (Get-Email -NTID "jsappi")
    List-Factors -email (Get-Email -First "Justin" -Last "Sappington")

    List-Factors -NTID "jsappi"
    List-Factors -First "Justin" -Last "Sappington"

    Add-Factor -SMS -Phone "+15032706690" -email "Justin.Sappington@nike.com"
    Add-Factor -Voice -Phone "+15032706690" -email "Justin.Sappington@nike.com"
    Add-Factor -Google -email "Justin.Sappington@nike.com"
    Add-Factor -email "Justin.Sappington@nike.com" -Phone "+15032706690" -Voice -SMS -Google
    Add-Factor -NTID "jsappi" -Google

    Remove-Factor -SMS -Email "justin.sappington@nike.com"
    Remove-Factor -Voice -Email "Justin.sappington@nike.com"
    Remove-Factor -Google -email "Justin.Sappington@nike.com"
    Remove-Factor -email "Justin.Sappington@nike.com" -Voice -SMS -Google

    Add-Factor-CSV -Google -Path "C:\Users\jsappi\Downloads\Wave 10million.csv" -email_Col_Name "Nike Email"
    Add-Factor-CSV -Google -Path "C:\Users\jsappi\Downloads\Wave 10million.csv" -NTID_Col_Name "User ID"
    Add-Factor-CSV -Google -Path "C:\Users\jsappi\Downloads\Wave 10million.csv" -NTID_Col_Name "User ID" -MFA_Col "MFA"
    Add-Factor-CSV -Google -Path "C:\Users\jsappi\Downloads\Wave 10million.csv" -FirstName_Col "First -LastName_Col "Last" `
        -MFA_Col "MFA Code" -NoForce
#>

#this imports all the dependencies
#these two lines are required for the Google TOTP generation
using namespace System
$Script:Base32Charset = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567'

#this module is needed if the email is not available (IE only have NTID or First/Last)
Import-Module ActiveDirectory
$config = Get-Content "$PSScriptRoot\Okta_API_Config.ini"

$headers = @{}
$baseUrl = $config[1].Substring($config[1].IndexOf("=")+2)
$userAgent = ""
$token = $config[2].Substring($config[2].IndexOf("=")+2)

#region Google Auth Pin
<#
.Synopsis
  Takes a Google Authenticator secret like 5WYYADYB5DK2BIOV
  and generates the PIN code for it
.Example
  PS C:\>Get-GoogleAuthenticatorPin -Secret 5WYYADYB5DK2BIOV
  372 251
.LINK
  Got this from the link below
  https://github.com/HumanEquivalentUnit/PowerShell-Misc/blob/master/GoogleAuthenticator.psm1
#>
function Get-GoogleAuthenticatorPin
{
    [CmdletBinding()]
    Param
    (
        # BASE32 encoded Secret e.g. 5WYYADYB5DK2BIOV
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $Secret,

        # OTP time window in seconds
        $TimeWindow = 30
    )


    # Convert the secret from BASE32 to a byte array
    # via a BigInteger so we can use its bit-shifting support,
    # instead of having to handle byte boundaries in code.
    $bigInteger = [Numerics.BigInteger]::Zero
    foreach ($char in ($secret.ToUpper() -replace '[^A-Z2-7]').GetEnumerator()) {
        $bigInteger = ($bigInteger -shl 5) -bor ($Script:Base32Charset.IndexOf($char))
    }

    [byte[]]$secretAsBytes = $bigInteger.ToByteArray()
    

    # BigInteger sometimes adds a 0 byte to the end,
    # if the positive number could be mistaken as a two's complement negative number.
    # If it happens, we need to remove it.
    if ($secretAsBytes[-1] -eq 0) {
        $secretAsBytes = $secretAsBytes[0..($secretAsBytes.Count - 2)]
    }


    # BigInteger stores bytes in Little-Endian order, 
    # but we need them in Big-Endian order.
    [array]::Reverse($secretAsBytes)
    

    # Unix epoch time in UTC and divide by the window time,
    # so the PIN won't change for that many seconds
    $epochTime = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    
    # Convert the time to a big-endian byte array
    $timeBytes = [BitConverter]::GetBytes([int64][math]::Floor($epochTime / $TimeWindow))
    if ([BitConverter]::IsLittleEndian) { 
        [array]::Reverse($timeBytes) 
    }

    # Do the HMAC calculation with the default SHA1
    # Google Authenticator app does support other hash algorithms, this code doesn't
    $hmacGen = [Security.Cryptography.HMACSHA1]::new($secretAsBytes)
    $hash = $hmacGen.ComputeHash($timeBytes)


    # The hash value is SHA1 size but we want a 6 digit PIN
    # the TOTP protocol has a calculation to do that
    #
    # Google Authenticator app may support other PIN lengths, this code doesn't
    
    # take half the last byte
    $offset = $hash[$hash.Length-1] -band 0xF

    # use it as an index into the hash bytes and take 4 bytes from there, #
    # big-endian needed
    $fourBytes = $hash[$offset..($offset+3)]
    if ([BitConverter]::IsLittleEndian) {
        [array]::Reverse($fourBytes)
    }

    # Remove the most significant bit
    $num = [BitConverter]::ToInt32($fourBytes, 0) -band 0x7FFFFFFF
    
    # remainder of dividing by 1M
    # pad to 6 digits with leading zero(s)
    # and put a space for nice readability
    $PIN = ($num % 1000000).ToString().PadLeft(6, '0').Insert(3, ' ')

    [PSCustomObject]@{
        'PIN Code' = $PIN
        'Seconds Remaining' = ($TimeWindow - ($epochTime % $TimeWindow))
    }
}

#endregion Google Auth Pin

#region Okta API Functions

#credit to https://github.com/mbegan/Okta-PSModule
#downloaded from https://github.com/gabrielsroka/OktaAPI.psm1

#region Core functions

# Call Connect-Okta before calling Okta API functions.
function Connect-Okta($token, $baseUrl) {
    $script:headers = @{"Authorization" = "SSWS $token"; "Accept" = "application/json"; "Content-Type" = "application/json"}
    $script:baseUrl = $baseUrl

    #$module = Get-Module OktaAPI
    $modVer = '1.0.16' #$module.Version.ToString()
    $psVer = $PSVersionTable.PSVersion

    $osDesc = [Runtime.InteropServices.RuntimeInformation]::OSDescription
    $osVer = [Environment]::OSVersion.Version.ToString()
    if ($osDesc -match "Windows") {
        $os = "Windows"
    } elseif ($osDesc -match "Linux") {
        $os = "Linux"
    } else { # "Darwin" ?
        $os = "MacOS"
    }

    $script:userAgent = "okta-api-powershell/$modVer powershell/$psVer $os/$osVer"
    # $script:userAgent = "OktaAPIWindowsPowerShell/0.1" # Old user agent.
    # default: "Mozilla/5.0 (Windows NT; Windows NT 6.3; en-US) WindowsPowerShell/5.1.14409.1012"

    # see https://www.codyhosterman.com/2016/06/force-the-invoke-restmethod-powershell-cmdlet-to-use-tls-1-2/
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

function Invoke-Method($method, $path, $body) {
    $url = $baseUrl + $path
    if ($body) {
        $jsonBody = $body | ConvertTo-Json -compress -depth 100 # max depth is 100. pipe works better than InputObject
        # from https://stackoverflow.com/questions/15290185/invoke-webrequest-issue-with-special-characters-in-json
        # $jsonBody = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)
    }
    Invoke-RestMethod $url -Method $method -Headers $headers -Body $jsonBody -UserAgent $userAgent
}

function Invoke-PagedMethod($url, $convert = $true) {
    if ($url -notMatch '^http') {$url = $baseUrl + $url}
    $response = Invoke-WebRequest $url -Method GET -Headers $headers -UserAgent $userAgent
    $links = @{}
    if ($response.Headers.Link) { # Some searches (eg List Users with Search) do not support pagination.
        foreach ($header in $response.Headers.Link.split(",")) {
            if ($header -match '<(.*)>; rel="(.*)"') {
                $links[$matches[2]] = $matches[1]
            }
        }
    }
    $objects = $null
    if ($convert) {
        $objects = ConvertFrom-Json $response.content
    }
    @{objects = $objects
      nextUrl = $links.next
      response = $response
      limitLimit = [int][string]$response.Headers.'X-Rate-Limit-Limit'
      limitRemaining = [int][string]$response.Headers.'X-Rate-Limit-Remaining' # how many calls are remaining
      limitReset = [int][string]$response.Headers.'X-Rate-Limit-Reset' # when limit will reset, see also [DateTimeOffset]::FromUnixTimeSeconds(limitReset)
    }
}

function Invoke-OktaWebRequest($method, $path, $body) {
    $url = $baseUrl + $path
    if ($body) {
        $jsonBody = $body | ConvertTo-Json -compress -depth 100
    }
    $response = Invoke-WebRequest $url -Method $method -Headers $headers -Body $jsonBody -UserAgent $userAgent
    @{objects = ConvertFrom-Json $response.content
      response = $response
      limitLimit = [int][string]$response.Headers.'X-Rate-Limit-Limit'
      limitRemaining = [int][string]$response.Headers.'X-Rate-Limit-Remaining' # how many calls are remaining
      limitReset = [int][string]$response.Headers.'X-Rate-Limit-Reset' # when limit will reset, see also [DateTimeOffset]::FromUnixTimeSeconds(limitReset)
    }
}

function Get-Error($_) {
    $responseStream = $_.Exception.Response.GetResponseStream()
    $responseReader = New-Object System.IO.StreamReader($responseStream)
    $responseContent = $responseReader.ReadToEnd()
    ConvertFrom-Json $responseContent
}
#endregion

#region Factors (MFA) - https://developer.okta.com/docs/api/resources/factors

function Get-OktaFactor($userid, $factorid) {
    Invoke-Method GET "/api/v1/users/$userid/factors/$factorid"
}

function Get-OktaFactors($userid) {
    Invoke-Method GET "/api/v1/users/$userid/factors"
}

function Get-OktaFactorsToEnroll($userid) {
    Invoke-Method GET "/api/v1/users/$userid/factors/catalog"
}

function Set-OktaFactor($userid, $factor, $activate = $false, [Switch]$phone) {
    If($phone){
        #without the updatePhone=true parameter, this cannot change the phone number and will error out
        #if the user has an activated phone setup already
        Invoke-Method POST "/api/v1/users/$userid/factors?updatePhone=true&activate=true" $factor
    }

    Else{
        Invoke-Method POST "/api/v1/users/$userid/factors?activate=$activate" $factor
    }
}

function Enable-OktaFactor($userid, $factorid, $body) {
    Invoke-Method POST "/api/v1/users/$userid/factors/$factorid/lifecycle/activate" $body
}

function Remove-OktaFactor($userid, $factorid) {
    $null = Invoke-Method DELETE "/api/v1/users/$userid/factors/$factorid"
}
#endregion

#region Okta User Functions (only relevant ones are present)

function Get-OktaUser($id) {
    Invoke-Method GET "/api/v1/users/$id"
}

function Set-OktaUser($id, $body) {
# Only the profile properties specified in the request will be modified when using the POST method.
    Invoke-Method POST "/api/v1/users/$id" $body
}

#endregion Okta User Functions

#endregion Okta API Functions

#region Nike Made Functions #JDI #swoosh #keepittight #MSADH

#using first and last is highly not reccommended as it may return ambiguous results (multiple users)
function Get-Email{
    Param(
        [Parameter(mandatory=$false)]$NTID="",
        [Parameter(mandatory=$false)]$First="",
        [Parameter(mandatory=$false)]$Last=""  
        )

    If($NTID -ne ""){
        $clean_name = $null
        $clean_name = $NTID -replace " ",""
        $clean_name = $clean_name -replace "`n",""
        Return ((Get-ADUser $clean_name | Select-Object userprincipalname).userprincipalname)
    }

    ElseIf($First -ne ""){
        Return ((Get-ADUser -f{(givenname -like $First) -And (surname -like $Last)}).userprincipalname)
    }
}

function List-Factors{
    Param(
        [Parameter(mandatory=$false)]$Email="",
        [Parameter(mandatory=$false)]$NTID="",
        [Parameter(mandatory=$false)]$First="",
        [Parameter(mandatory=$false)]$Last="",
        [Switch]$No_Connect       
    )
    
    #need to determine what to do based on the input
    If($NTID -ne ""){
        #call Get-Email with NTID
        $email = Get-Email -NTID $NTID
    }

    ElseIf($First -ne ""){
        If($Last -eq ""){
            Write-Output "Please provide First and Last name"
            Break
        }
        #call Get-Email with first and last
        $email = Get-Email -First $First -Last $Last
    }

    ElseIf($Last -ne ""){
        Write-Output "Please provide First and Last name."
        Break    
    }

    ElseIf($Email -eq ""){
        Write-Output "Please provide input."
        Break
    }
    
    #we have the goods, so connect to Okta first
    If($No_Connect -ne $true){Connect-Okta $token $baseUrl}
    #get okta userID
    Try{
        $id = $null
        $id = (Get-OktaUser $Email).id
        #Get the okta factors
        Get-OktaFactors $id
    }
    Catch{
        If($_ -like "*(404) Not Found*"){Return "User not found."}
        Else{Return $_}
        Break
    }
}

function Add-Factor{
    Param(
        [Parameter(mandatory=$false)]$Email="",
        [Switch]$SMS,
        [Switch]$Voice,
        [Switch]$Google,
        [Parameter(mandatory=$false)]$Phone="",
        [Switch]$No_Connect,
        [Parameter(mandatory=$false)]$NTID=""
    )
    #we have the goods, so connect to Okta first, unless told otherwise
    #that is for Add-Factor-CSV so we don't hit rate limits
    If($No_Connect -ne $true){Connect-Okta $token $baseUrl}
    
    If($NTID -ne ""){
        #call Get-Email with NTID
        $Email = Get-Email -NTID $NTID
    }
    Elseif($Email -eq ""){
        Write-Output "Please provide user email or NTID."
        Break
    }

    If(($SMS -or $Voice) -and $Phone -eq ""){
        Write-Output "Please provide a phone number"
        Break
    }

    Try{
        $id = $null
        $id = (Get-OktaUser $Email).id
        $available_factors = $null
        $available_factors =  Get-OktaFactorsToEnroll $id
    }
    Catch{
        If($_ -like "*(404) Not Found*"){Return "User not found."}
        Else{Return $_}
        Break
    }

    $count = 0

    If($SMS){        
        $new_SMS_factor = $null
        $new_SMS_factor = @{factorType = "sms"; provider = "OKTA"; profile = @{phoneNumber = $Phone}}
        $activate = $true # Activate SMS without sending one to the user.
        Try{
            $new_one = $null; $new_one = Set-OktaFactor $id $new_SMS_factor $activate -phone
            $output = ""; $output+= $new_one.factorType + ",  " + $new_one.provider + ", " + $new_one.profile.phoneNumber + ", " + $new_one.status
            Write-Output $output
            $count++
        }
        Catch{
            $error = $null
            $error = Get-Error $_
            #invalid phone number error code
            If($error.errorCode -eq "E0000098"){
                Return $error.errorSummary + " (SMS)"
            }
            Elseif($error.errorCode -eq "E0000001"){                
                Return $error.errorCauses.errorSummary + " (SMS)"
            }
            Else{return $error}
        }
    }

    If($Voice){
        $new_Voice_factor =$null
        $new_Voice_factor = @{factorType = "call"; provider = "OKTA"; profile = @{phoneNumber = $Phone}}
        $activate = $true # Activate Voice without calling the user.
        Try{
            $new_one = $null; $new_one = Set-OktaFactor $id $new_Voice_factor $activate -phone
            $output = ""; $output+= $new_one.factorType + ", " + $new_one.provider + ", " + $new_one.profile.phoneNumber + ", " + $new_one.status
            Write-Output $output
            $count++
        }
        Catch{(Get-Error $_).errorCauses.errorSummary + " (Voice)"}
    }

    If($Google){
        $new_google_factor = $null
        $new_google_factor = @{factorType = 'token:software:totp'; provider = 'GOOGLE'}
        Try{
            #this block generates the 16 digit activation code, but doesn't activate the factor for the end user
            $new_goog = $null
            $new_goog = Set-OktaFactor $id $new_google_factor
            $new_goog_act_code = $new_goog._embedded.activation.sharedSecret
            
            #now we need to activate the factor
            #This involves getting a 6 digit pin and giving it to okta to confirm we have the right secret
            #Get-GoogleAuthenticatorPin -secret $new_goog._embedded.activation.sharedSecret
            #we also need to handle the edge case where the TOTP password expires before we can use it with Okta
            #if the TOTP expires in the next 3s, we will have the script sleep for the time until it expires +1s
            #so that we don't have issues with failing to setup because of something simple
            $goog_auth_result = $null
            $goog_auth_result = Get-GoogleAuthenticatorPin -secret $new_goog._embedded.activation.sharedSecret
            
            #here we check if the TOTP will expire soon and then sleep if it does,
            #generating a new TOTP when it wakes up
            #Two seconds was too short and still had the same issue. Set it to 3
            If($goog_auth_result.'Seconds Remaining' -le 3){
                Start-Sleep -s ($goog_auth_result.'Seconds Remaining' +1)
                $goog_auth_result = Get-GoogleAuthenticatorPin -secret $new_goog._embedded.activation.sharedSecret
            }

            $TOTP = $null
            #format the TOTP as okta didn't like the space between the two trios
            #111 222 becomes 111222           
            $TOTP = $goog_auth_result.'PIN Code' -replace " ", ""
            $TOTP_son = @{passCode = $TOTP}            
            $new_goog_active = $null
            $new_goog_active = Enable-OktaFactor $id $new_goog.id $TOTP_son
            
            #write out new MFA Activation Code
            $output = ""; $output+= $new_goog_active.Provider + ", " + "$Email - $new_goog_act_code" + ", " + $new_goog_active.status
            Write-Output $output
            $count++
        }
        Catch{(Get-Error $_).errorCauses.errorSummary + " (Google)"}
        
    }
    If($count -eq 1){Write-Output "`nSetup $count factor for user $Email"}
    Else{Write-Output "`nSetup $count factors for user $Email"}
}

function Remove-Factor{
    Param(
        [Parameter(mandatory=$false)]$Email="",
        [Parameter(mandatory=$false)]$NTID="",
        [Switch]$SMS,
        [Switch]$Voice,
        [Switch]$Google,
        [Switch]$No_Connect
    )

    #we have the goods, so connect to Okta first, unless told otherwise
    #that is for Add-Factor-CSV so we don't hit rate limits
    If($No_Connect -ne $true){Connect-Okta $token $baseUrl}

    If($NTID -ne ""){
        #call Get-Email with NTID
        $Email = Get-Email -NTID $NTID
    }
    Elseif($Email -eq ""){
        Write-Output "Please provide user email or NTID."
        Break
    }

    #now we need the Factor IDs as Okta is all about dem IDs
    $factor_list = $null
    Try{$factor_list = List-Factors -Email $Email}
    Catch{
        If($_ -like "*(404) Not Found*"){Return "User not found."}
        Else{$_}
        Break
    }
    $sms_fac_id = $null
    $voice_fac_id = $null
    $Google_fac_id = $null
    $count = 0

    $id = $null
    Try{$id = (Get-OktaUser $Email).id}
    Catch{
        If($_ -like "*(404) Not Found*"){Return "User not found."}
        Else{
            $_
            Break
        }
        
    }

    #loop through the list of factors and set the IDs
    Foreach($line in $factor_list){
        If($line.FactorType -eq "call"){$voice_fac_id = $line.id}
        ElseIf($line.FactorType -eq "sms"){$sms_fac_id = $line.id}
        ElseIf($line.provider -eq "GOOGLE"){$Google_fac_id = $line.id}
    }

    If($SMS){
        Try{Remove-OktaFactor $id $sms_fac_id;$count++}
        Catch{
            $error = $null
            $error = Get-Error $_
            If($error.errorCode -like "*E0000022*"){Write-Output "SMS Factor is not setup."}
            Else{Return $error.ErrorSummary}
        }
    }

    If($Voice){
        Try{Remove-OktaFactor $id $voice_fac_id;$count++}
        Catch{
            $error = $null
            $error = Get-Error $_
            If($error.errorCode -like "*E0000022*"){Write-Output "Voice Factor is not setup."}
            Else{Return $error.ErrorSummary}
        }
    }

    If($Google){
        Try{Remove-OktaFactor $id $Google_fac_id;;$count++}
        Catch{
            $error = $null
            $error = Get-Error $_
            If($error.errorCode -like "*E0000022*"){Write-Output "Google Factor is not setup."}
            Else{Return $error.ErrorSummary}
        }
    }

    If($count -eq 1){Write-Output "Removed $count factor for user $Email"}
    Else{Write-Output "Removed $count factors for user $Email"}
}

function Add-Factor-CSV{
    Param(
        [Parameter(mandatory=$true)]$Path="",
        #not implementing these yet
        #[Switch]$SMS,
        #[Switch]$Voice,
        [Switch]$Google,
        [Switch]$No_Connect,
        [Parameter(mandatory=$false)]$NTID_Col_Name="",
        [Parameter(mandatory=$false)]$Email_Col_Name="",
        [Parameter(mandatory=$false)]$FirstName_Col="",
        [Parameter(mandatory=$false)]$LastName_Col="",
        [Parameter(mandatory=$false)]$MFA_Col="",
        [Switch]$No_Force
    )

    $in_csv = $null
    $in_csv = Import-Csv $path

    Connect-Okta $token $baseUrl
    
    #testing outputs
    #$in_csv
    #$email_Col_Name
    #$NTID_Col_Name

    #if all column name parameters are empty, exit and tell em to give us column names
    if(($email_Col_name -eq "") -and `
        ($NTID_Col_Name -eq "") -and `
        ($FirstName_Col -eq "") -and `
        ($LastName_Col  -eq "")){
            Write-Output "`nPlease provide any of the following options:`n"
            Write-Output "First and Last Name column names,`nNTID column name,`nEmail column name."
            Return
        }

    #if email_col is blank, create that column
    if($email_Col_name -eq ""){
        $email_Col_Name = "User Email"
        $in_csv = $in_csv | Select-Object *, @{name = $email_Col_Name;expr={}}
    }

    #if MFA_col is blank, search for columns -like "*MFA*".  If found, use that column
    #for the MFA Activation code output.  Otherwise, create that column with the name "MFA Activation Code"
    if($MFA_Col -eq ""){
        $column_names = $null
        $column_names = (Get-Member -InputObject $in_csv[0]).name
        Foreach($object in $column_names){
            If($object -like "*MFA*"){
                $MFA_Col = $object
                Break
            }
        }

        #no MFA column exists, so create one with the name MFA Activation Code
        if($MFA_Col -eq ""){
            $MFA_Col = "MFA Activation Code"
            $in_csv = $in_csv | Select-Object *, @{name = $MFA_Col;expr={}}
        }
    }

    foreach($user in $in_csv){
        $clean_name = $null
        $goog = $null

        Try{
            #if we don't have the email, we need to get the email.  Either via NTID or First/last names
            if($user.$email_Col_Name -eq $null){
                #first check for NTID
                if($user.$NTID_Col_Name -ne ""){
                    $clean_name = $user.$NTID_Col_Name -replace " ",""
                    $clean_name = $clean_name -replace "`n",""
                    $user.$email_Col_Name = (Get-Email -NTID $clean_name)
                }
                #then check to see if we have first/last name
                elseif($FirstName_Col -eq ""){
                    Write-Output "No user data provided."
                    Break
                }
                Elseif($LastName_Col -eq ""){
                    Write-Output "Please provide first and last name."
                    Break
                }
                Else{
                    $user.$email_Col_Name = Get-Email -First $user.$FirstName_Col -Last $user.$LastName_Col
                }
            }
            
            #need to clean the email of any whitespace
            $user.$email_Col_Name = $user.$email_Col_Name -replace " ",""
            $user.$email_Col_Name = $user.$email_Col_Name -replace "`n",""
                        
            #if the $no_force flag is set, then we do not attempt to delete any old google mfa
            if($No_Force -ne $true){
                $remove_result = $null
                $remove_result = Remove-Factor -Email $user.$email_Col_Name -Google -No_Connect

                If($remove_result -eq "User not found."){
                    $user.$MFA_Col = $remove_result
                    $out_text = $null
                    $out_text = ($user.$email_Col_Name)+" - "+($user.$MFA_Col)
                    Write-Output $out_text
                    Continue
                }
                Write-Output $remove_result
            }

            #now we add the new factor and write it to the output object
            $goog = Add-Factor -Google -Email $user.$email_Col_Name -No_Connect

            if($goog -like "*A factor of this type is already set up.*"){
                $user.$MFA_Col = $goog[0]#"A factor of this type is already set up."
            }
            Elseif($goog -eq "User not found."){
                $user.$MFA_Col = "User not found."
            }
            else{
                $user.$MFA_Col = $goog[0].Substring($goog[0].LastIndexOf("-")+2,16)
            }
            
        }
        Catch{
            If($_ -like "*(404) Not Found*"){$user.$MFA_Col = "User not found."}
            Else{$_}
        }
        $out_text = $null
        $out_text = ($user.$email_Col_Name)+" - "+($user.$MFA_Col)
        Write-Output $out_text
    }

    #won't overwrite if CSV is open which could result in data loss, need to create a new file
    #generate export file name
    $new_csv_path = $path.substring(0,$path.lastindexof("\")+1)
    #get the date for adding to file name
    $dtg = Get-Date -Format "MM-dd-yyyy_HHmmss"
    #create file name
    $new_csv_file_name = $new_csv_path+"MFA Activation Codes_"+$dtg+".csv"
    $in_csv | Export-Csv -Path $new_csv_file_name -NoTypeInformation -Force
}

#endregion