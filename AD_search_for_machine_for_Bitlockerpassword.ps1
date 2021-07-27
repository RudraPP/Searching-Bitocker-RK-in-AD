<#

Author : Rudra Prasad Paul
#>

[CmdletBinding()]
Param (
    [string]$OUName = "OU=MBAM,DC=variablex,DC=com"
)

$Export = @()

 

$ExportCSV = $false
 
$choice = Read-Host "Enter your choice `n1. Run against OU where the OU is -> $OUName`n2. Run against entire Directory(all Machines)`n`nSelect from the above. Enter 1 or 2 "


if($choice -eq "1")
{
    $csvChoice = Read-Host "`nNeed a CSV file as output? Type Y or N "
    if($csvChoice -eq "Y")
    {
        $ExportCSV=$true
    }

    $machinename = Get-ADComputer -Filter * -SearchBase $OUName -Properties *
     Write-Host "`nSearching through OU : ",$OUName
}


elseif($choice -eq "2")
{
    $csvChoice = Read-Host "`nNeed a CSV file as output? Type Y or N "
    if($csvChoice -eq "Y")
    {
        $ExportCSV=$true
    }

    $machinename = Get-ADComputer -Filter *  -Properties *
     Write-Host "`nSearching through all the machines in AD"
}

else
{
    Write-Host "`n`nSelected option doesn't exit......`n`nEnding script"
    break
}

   
 
 

     foreach ($machine in $machinename)
    {
        $MN = $machine.Name
        $bitlockerpassword = Get-ADObject -Filter "objectClass -eq 'msFVE-RecoveryInformation'" -SearchBase $machine.distinguishedName -Properties msFVE-RecoveryPassword,whenCreated | Sort whenCreated -Descending | Select -First 1 | Select -ExpandProperty msFVE-RecoveryPassword
        $createdon = Get-ADObject -Filter "objectClass -eq 'msFVE-RecoveryInformation'" -SearchBase $machine.distinguishedName -Properties msFVE-RecoveryPassword,whenCreated | Sort whenCreated -Descending | Select -First 1 | Select -ExpandProperty whencreated
    
        $obj = New-Object -TypeName psobject
        $obj | Add-Member -MemberType NoteProperty -Name MachineName -Value "$MN"
               
        If ($bitlockerpassword)
        {
            $obj | Add-Member -MemberType NoteProperty -Name 'BitLocker Password' -Value $bitlockerpassword
            $obj | Add-Member -MemberType NoteProperty -Name 'Created on' -Value $createdon
        }
        else
        {
            $obj | Add-Member -MemberType NoteProperty -Name 'BitLocker Password' -Value "Not Present"
            $obj | Add-Member -MemberType NoteProperty -Name 'Created on' -Value "----"
        }

 

        $obj | Format-List 
        $Export+=$obj

   
    }






    if ($ExportCSV -eq $true){

    
        $Export | Export-Csv -Path c:\temp\BitlockerRecoveryexport.csv -NoTypeInformation
        Write-Host "`nCSV file is stored in c:\temp\BitlockerRecoveryexport.csv"

    }