#####################################################################################
#
# Get-Parsed_Nessus.psm1
# This function parses Nessus XMLs and updates global hashtable (Yeah ... I know)
#
# Parameters Modified: 
# $IssueKey 
#
# Author: https://github.com/CBHue/
#  
# Version: 1.0
#
#####################################################################################

function Get-Parsed_Nessus {
    if ((get-pscallstack |select -last 2 |select -expa arguments -first 1) -match "debug"){ $debugpreference="continue"}
    write-output "Welcome to the Nessus Parser ... $($file.name)"
    $xmldoc = new-object System.XML.XMLdocument
        
    # Some Nessus files have issues 
    Try{ [xml]$xmldoc = Get-Content $file -ReadCount 0 }
    #Try{ $xmldoc.Load($file) }
    Catch [system.exception] {
        write-output "We had an issue opening so were converting the file ..."
        # Some house keeping to replace troublesome XML characters
        Copy-Item $file "$file.BAK"
        (Get-Content $file) | 
        Foreach-Object {$_ -replace "&", "&amp;" -replace "'", "&apos;" -replace "–", " " -replace "(?<!\?xml.*)(?<=`".*?)`"(?=.*?`")", ""} |
        Set-Content $file
        [xml]$xmldoc = Get-Content $file -ReadCount 0
        #$xmldoc.Load($file)
    }

    $vID = $IssuevID.count
    $vID++
    
    $reportHost = $xmldoc.NessusClientData_v2.Report         
    $hst = $reportHost.ReportHost.count
    
    # If there is only one host its not an array    
    if ($hst -eq $null) {$hst = 1 }
    Write-Debug "+ $hst hosts to parse ..."
    $hst--
    
    foreach ($j in (0..$hst)) {
            if ($hst -ne 0) {$reportHost = $xmldoc.NessusClientData_v2.Report.ReportHost[$j]}
            else            {$reportHost = $xmldoc.NessusClientData_v2.Report.ReportHost}
            
            $hostID = $reportHost.name
            $vCount = $reportHost.reportitem.count
            $vCount--
                
            # Ok now we got all the info we need to loop over the hosts and for each host loop the findings 
            $version = $reportHost.ReportItem[0].plugin_output | select-string Nessus | % {$_.line.split()[12]}

            foreach ($i in (0..$vCount)) {

                # Lets update the report host
                $reportItem = $reportHost.ReportItem[$i]
                
                # If this is general info move on
                if ($reportItem."svc_name" -eq "general"){ continue }

                # build the key ip - port       
                $sb = New-Object System.Text.StringBuilder
                $null = $sb.Append("$($hostID)-$($reportItem."port")")
                $name = $sb.tostring().Trim()

                # we need a unique Key to save our info ... im hashing the desciption value and using it
                #$hName = Hash("$name")
                $hName = $name

                # Skip if we already have this host - ip key
                if ($IssueKey.ContainsKey($hName)){ continue }
                $svc = ($reportItem."svc_name").trim("?")

                $IssueKey.Add(($hName),(@{
                    "ip"         = ($hostID)
                    "port"       = ($reportItem."port")
                    "protocol"   = ($reportItem."protocol")
                    "svc_name"   = ($svc.ToLower())
                }))
               
                # update our Key ID
                $vID++
            } # End vLoop

            Write-Debug "+ Finsished $j of $hst : $hostID "
        } # End Host loop 
}

function Hash($t) {
    $md5 = new-object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    $utf8 = new-object -TypeName System.Text.UTF8Encoding
    $hash = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($t)))
    return $hash;
}

#
# End: Get-Parsed_Nessus.psm1
#