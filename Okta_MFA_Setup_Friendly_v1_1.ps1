<#
.SYNOPSIS
    This is the friendly script for administrating MFA in Okta.  It will ask questions and call the
    functions needed based on your answers.
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
            If they are not given, the script will create them)
.NOTES
    Sources : https://github.com/gabrielsroka/OktaAPI.psm1
              https://github.com/HumanEquivalentUnit/PowerShell-Misc/blob/master/GoogleAuthenticator.psm1
    Author  : Justin Sappington
    Created : October 28th, 2019
    Last Modified : October 31st, 2019
    Dependencies  : Okta_MFA_Setup.ps1
                  : Powershell ActiveDirectory Module
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
        Date: 10-31-2019 - Actor: Justin Sappington
        Description:
                    -Cleaned up all the function calls below to make them more user friendly.
                    Reorganized them and put comments for how to use them above the calls
        Date: 02-05-2020 - Actor: Justin Sappington
        Description:
                    -Adding front end support for removing HOTP factors from Okta users.
                    -Added prompt option: "5. Hardware MFA Token (HOTP)" to "What factor(s) do you want to modify?"
                    prompt
                    -Added  ElseIf($what_factor_choice -eq 5){
                                Write-Output "Setting up HOTP factors is self service in Okta.  It is not currently supported by this solution."
                            }
                    to add factor choice.
                    -Added      ElseIf($what_factor_choice -eq 5){
                                    #email
                                    If($what_have_choice -eq 1){Remove-Factor -Email $User_Email -HOTP}
                                    #NTID
                                    Else{Remove-Factor -NTID $User_NTID -HOTP}
                                }
                    to the remove factor choice.


.EXAMPLE
    Press F5 to run this script. It will look similar to below.
    
    What do you want to do?
    1. Add Factor(s)
    2. Remove Factor(s)
    3. List Factors that are currently setup.
    4. Add google factors from a CSV
    Choice: 2


    What user info do you have?
    1. Email
    2. Sam Account Name/LogonID
    Choice: 2
    Sam Account Name/LogonID: jsappi


    What factor(s) do you want to modify?
    1. SMS
    2. Voice
    3. Google
    4. SMS and Voice
    Choice: 1
#>

Import-Module "$PSScriptRoot\Okta_MFA_Setup_v1_0.ps1" -Force

#region What do
Write-Output "What do you want to do?","1. Add Factor(s)","2. Remove Factor(s)",`
"3. List Factors that are currently setup.", "4. Add Google factors from a CSV"
$what_do_choice = Read-Host -Prompt "Choice"
"`n"
#endregion

#region What got
If($what_do_choice -in 1,2,3){
    Write-Output "What user info do you have?", "1. Email", "2. Sam Account Name/LogonID"
    $what_have_choice = Read-Host -Prompt "Choice"

    #email
    If($what_have_choice -eq 1){
        $User_Email = Read-Host -Prompt "Email"        
    }
    #NTID/Sam Account Name
    Elseif($what_have_choice -eq 2){
        $User_NTID = Read-Host -Prompt "Sam Account Name/LogonID"
    }
    "`n"
}
#faulty input
Elseif($what_do_choice -eq 4){
    Write-Output "`nWhat user info do you have?", "1. Email", "2. Sam Account Name/LogonID"
    $what_have_choice = Read-Host -Prompt "Choice"
}
Else{
    Write-Output "Bad input.  Just enter the number of your choice"
    Break
}

#endregion

#region What factor do

#endregion

#region IFFFFFSSSSSS

#add factor or remove factor
If(($what_do_choice -eq 1) -or ($what_do_choice -eq 2)){
    Write-Output "What factor(s) do you want to modify?","1. SMS","2. Voice","3. Google", "4. SMS and Voice", "5. Hardware MFA Token (HOTP)"
    $what_factor_choice = Read-Host -Prompt "Choice"
    "`n"

    #add-factor
    If($what_do_choice -eq 1){
        #get phone number
        If(($what_factor_choice -eq 1) -or ($what_factor_choice -eq 2) -or ($what_factor_choice -eq 4)){
            Write-Output "Enter the phone number to setup for MFA"
            $User_Phone = Read-Host -Prompt "Phone Number"
            "`n"
        }
        #SMS
        If($what_factor_choice -eq 1){
            #email
            If($what_have_choice -eq 1){Add-Factor -Email $User_Email -Phone $User_Phone -SMS}
            #NTID
            Else{Add-Factor -NTID $User_NTID -Phone $User_Phone -SMS}
        }
        #Voice
        Elseif($what_factor_choice -eq 2){
            #email
            If($what_have_choice -eq 1){Add-Factor -Email $User_Email -Phone $User_Phone -Voice}
            #NTID
            Else{Add-Factor -NTID $User_NTID -Phone $User_Phone -Voice}
        }
        #google
        Elseif($what_factor_choice -eq 3){
            #email
            If($what_have_choice -eq 1){Add-Factor -Email $User_Email -Google}
            #NTID
            Else{Add-Factor -NTID $User_NTID -Google}
        }
        ElseIf($what_factor_choice -eq 4){
            #email
            If($what_have_choice -eq 1){Add-Factor -Email $User_Email -Phone $User_Phone -Voice -SMS}
            #NTID
            Else{Add-Factor -NTID $User_NTID -Phone $User_Phone -Voice -SMS}
        }
        ElseIf($what_factor_choice -eq 5){
            Write-Output "Setting up HOTP factors is self service in Okta.  It is not currently supported by this solution."
        }
    }

    #remove-factor
    ElseIf($what_do_choice -eq 2){
        #SMS
        If($what_factor_choice -eq 1){
            #email
            If($what_have_choice -eq 1){Remove-Factor -Email $User_Email -SMS}
            #NTID
            Else{Remove-Factor -NTID $User_NTID -SMS}
        }
        #Voice
        Elseif($what_factor_choice -eq 2){
            #email
            If($what_have_choice -eq 1){Remove-Factor -Email $User_Email -Voice}
            #NTID
            Else{Remove-Factor -NTID $User_NTID -Voice}
        }
        #google
        Elseif($what_factor_choice -eq 3){
            #email
            If($what_have_choice -eq 1){Remove-Factor -Email $User_Email -Google}
            #NTID
            Else{Remove-Factor -NTID $User_NTID -Google}
        }
        ElseIf($what_factor_choice -eq 4){
            #email
            If($what_have_choice -eq 1){Remove-Factor -Email $User_Email -Voice -SMS}
            #NTID
            Else{Remove-Factor -NTID $User_NTID -Voice -SMS}
        }
        ElseIf($what_factor_choice -eq 5){
            #email
            If($what_have_choice -eq 1){Remove-Factor -Email $User_Email -HOTP}
            #NTID
            Else{Remove-Factor -NTID $User_NTID -HOTP}
        }
    }

}

#List factors
Elseif($what_do_choice -eq 3){
    #email
    If($what_have_choice -eq 1){
        List-Factors -Email $User_Email
    }
    #NTID/Sam Account Name
    Elseif($what_have_choice -eq 2){
        #Get-ADUser $User_NTID
        List-Factors -NTID $User_NTID
    }
}

#Add factor csv
Elseif($what_do_choice -eq 4){
    #Input needed: file name, column names
    $location = Read-Host -Prompt "Please provide the full path and file name of the CSV`nExample: C:\Users\Jsappi\Downloads\Copy of MFA Wave 68-1.csv`nPath"
    "`n"
    #email
    If($what_have_choice -eq 1){
        $email_Column = Read-Host -Prompt "Email Column Name"
        "`n"
        Add-Factor-CSV -Path $location -Google -Email_Col_Name $email_Column
    }
    #NTID/Sam Account Name
    Elseif($what_have_choice -eq 2){
        $NTID_column = Read-Host -Prompt "Sam Account Name/LogonID Column Name"
        "`n"
        Add-Factor-CSV -Path $location -Google -NTID_Col_Name $NTID_column
    }
    
}
Else{
    Write-Output "Bad input.  Just enter the number of your choice"
}

#endregion
