####################################################################################
# 
# get-iplist.ps1 
# Parse XML output and create iplist in XLSX form 
#
# Description 
#    Convert XML output to XLSX reports
#
# Example 
#	 .\get-iplist.ps1 -path findingsFile.xml
#    .\get-iplist.ps1 -path *.xml 
#
# 
# Author: https://github.com/CBHue/
#
#   
# Supported XML: AppScan, Nessus, BurpSuite, webInspect, AppDetective
#
####################################################################################

[CmdletBinding()]
Param ($Path, [String] $OutputDelimiter = "`n")
$DebugPreference = "Continue"

# Version
$Version = "v1.6"

Function Get-FileName($initialDirectory) {   
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowHelp = $true
    $OpenFileDialog.Multiselect = $true
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.Filenames
    
} #end function Get-FileName
       
function create-iplist {
    write-output "$($sw.Elapsed) This is the Beginning ... $version"
    
    if ( $Path -eq $null ) { $Path = @(); $input | foreach { $Path += $_ } } 
    # To support wildcards in $path. 
    if ( ($Path -ne $null) -and ( $Path.gettype().name -eq "String")) {$Path = dir $path} 
    # To open File Dialog
    if ($($Path.length) -lt 1) { $Path = Get-FileName -initialDirectory "c:\" }
    
    #
    # Supported Parsers and theyre indicators
    #
    $parsER = @{"AppScan"          = "appscan"; 
                "Nessus"           = "NessusClientData_v2"; 
                "BurpSuite"        = "burpVersion"; 
                "WebInspect_Full"  = "<Sessions><Session"; 
                "webInspect_vOnly" = "<Scan><Name>"; 
                "AppDetective"     = "<CheckResults>"
                }

    # We need to save the information and write after we checked everyfile.
    # Powershell v3 ruined my Object hashes ... oh well back to the hash below
    #$global:IssueKey    = [PSCustomObject]  @{}
    $global:IssueKey      = @{}

    Foreach ($global:file in $Path) {
       if ($file -eq $null) {continue}
       # We need to figure out what kind of file it is and then call the associated parser.
       $head = Get-Content literalPath $file -totalcount 10
     
       # Loop thru the parsers to find what kind of file were loooking at
       $noMatch = 0
       $parsER.GetEnumerator() | Foreach-Object { 
            if ( $head | select-string $_.value -quiet) {
                $noMatch++
                $end   = "Finished with the " + $_.key + " Parser ..."
                $modR  = "Get-Parsed_" + $_.key
                $modI  = ".\Modules\" + $modR + ".psm1"

                # I should check to make sure the file is there ... maybe later
                if ((Get-Module $modR)) {
                    Write-Debug "Removing $modR"
                    Remove-Module $modR
                }
               
                Write-Debug "Importing $modI"
                Import-Module $modI
                Invoke-Expression "$modR -Debug:$DebugPreference"
                Write-Debug "$($sw.Elapsed) $end"
                Remove-Module $modR
            }           
       }
        
       if ($noMatch -eq 0) { 
           Write-Debug "File not recognized ... skipping $file" 
           continue        
       } 
           
    } # End of All Files
    Write-Debug "$($sw.Elapsed) Finished with all files"

    # Create the Output File ... Import current XLSX writer
    if ((Get-Module "XLSX writer")) { Remove-Module Set-XLSX_Output }

    Import-Module .\Modules\Set-XLSX_Output.psm1 -DisableNameChecking
    Write-Debug  "$($sw.Elapsed) Create XLSX Document"
    Create-XLSX
    Write-Debug  "$($sw.Elapsed) Inserting XLSX Header"
    Insert-Header
    Write-Debug  "$($sw.Elapsed) Inserting XLSX Body"
    Insert-Body
    Write-Debug  "$($sw.Elapsed) Formatting XLSX"
    Format-XLSX
    Write-output "$($sw.Elapsed) This is the END ..."
    Remove-Module Set-XLSX_Output 
} # End of Function

# Build hashtable for splatting the parameters:
$ParamArgs = @{ Path = $Path ; OutputDelimiter = $OutputDelimiter } 

# Allow XML files to be piped into script:
if ($ParamArgs.Path -eq $null) { $ParamArgs.Path = @(); $input | foreach { $ParamArgs.Path += $_ } } 

$global:sw = [Diagnostics.Stopwatch]::StartNew()

# Run the main function with the splatted params:
create-iplist @ParamArgs

$sw.Stop()
