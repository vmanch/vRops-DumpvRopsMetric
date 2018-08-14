#Powershell collector script for vRops Suite-api
<#

    .SYNOPSIS

    Collecting metric / state data from vROPS Suite-api and output's to CSV, run the script without parameters for instructions.

    Script requires Powershell v3 and above.

    Run the command below to store user and pass in secure credential XML for each environment

        $cred = Get-Credential
        $cred | Export-Clixml -Path "d:\vRops\config\HOME.xml"

#>

param
(
    [String]$vRopsAddress,
    [String]$CollectionType,
    [String]$creds,
    [Array]$ResourceByName,
    [Array]$Metrics,
    [String]$rollUpType,
    [String]$intervalType,
    [DateTime]$StartDate = (Get-date).adddays(-1),
    [DateTime]$EndDate = (Get-date),
    [String]$Format
)


#Get Stored Credentials

$ScriptPath = (Get-Item -Path ".\" -Verbose).FullName

if($creds -gt ""){

    $cred = Import-Clixml -Path "$ScriptPath\config\$creds.xml"

    $vRopsUser = $cred.GetNetworkCredential().Username
    $vRopsPassword = $cred.GetNetworkCredential().Password
    }
    else
    {
    echo "vRops creds missing, stop hammer time!"
    Exit
    }

#vars
$RunDateTime = (Get-date)
$RunDateTime = $RunDateTime.tostring("yyyyMMddHHmmss") 
$Share = '\\localhost\completed\'
$Output = $ScriptPath  + '\collections\' + $RunDateTime + '\'
New-Item $Output -type directory
$LogFileLoc = $ScriptPath + '\Log\Logfile.log'
$StartDateFile = $StartDate.tostring("yyyyMMdd-HHmmss")            
$EndDateFile = $EndDate.tostring("yyyyMMdd-HHmmss")

#JobControl used to limit how hard the machine's CPU and memory is used.
$maxJobCount = 8
$sleepTimer = 3

#Lookup ResourceId from Name

Function GetObject([String]$vRopsObjName, [String]$vRopsServer, $User, $Password){

$wc = new-object system.net.WebClient
$wc.Credentials = new-object System.Net.NetworkCredential($User, $Password)
[xml]$Checker = $wc.DownloadString("https://$vRopsServer/suite-api/api/resources?name=$vRopsObjName")

$AlertReport = @()

# Check if we get more than 1 result and apply some logic
    If ([Int]$Checker.resources.pageInfo.totalCount -gt '1') {

        $DataReceivingCount = $Checker.resources.resource.resourceStatusStates.resourceStatusState.resourceStatus -eq 'DATA_RECEIVING'

            If ($DataReceivingCount.count -gt 1){
            $CheckerOutput = ''
            return $CheckerOutput 
            }
            
            Else 
            {

            ForEach ($Result in $Checker.resources.resource){

                IF ($Result.resourceStatusStates.resourceStatusState.resourceStatus -eq 'DATA_RECEIVING'){

                     $PropertiesLink = $Result.links.link | where Name -eq 'latestPropertiesOfResource'
                     $Propertiesurl = 'https://' +$vRopsServer + $PropertiesLink.href
                     [xml]$Properties = $wc.DownloadString($Propertiesurl)

                     switch($Result.resourceKey.resourceKindKey)
                        {

                        VirtualMachine {

                            $ParentvCenter = $Properties.'resource-property'.property | where name -eq 'summary|parentVcenter' | Select '#text'
                            $ParentCluster = $Properties.'resource-property'.property | where name -eq 'summary|parentCluster' | Select '#text'
                            $ParentHost = $Properties.'resource-property'.property | where name -eq 'summary|parentHost' | Select '#text'
                            $PowerState = $Properties.'resource-property'.property | where name -eq 'summary|runtime|powerState' | Select '#text'
                            $Memory = $Properties.'resource-property'.property | where name -eq 'config|hardware|memoryKB' | Select '#text'
                            $CPU = $Properties.'resource-property'.property | where name -eq 'config|hardware|numCpu' | Select '#text'
                            $INFO = $Properties.'resource-property'.property | where name -eq 'config|guestFullName' | Select '#text'

                            }


                        HostSystem {

                            $ParentvCenter = $Properties.'resource-property'.property | where name -eq 'summary|parentVcenter' | Select '#text'
                            $ParentCluster = $Properties.'resource-property'.property | where name -eq 'summary|parentCluster' | Select '#text'
                            $ParentHost = $Properties.'resource-property'.property | where name -eq 'summary|parentHost' | Select '#text'
                            $PowerState = $Properties.'resource-property'.property | where name -eq 'runtime|powerState' | Select '#text'
                            $Memory = $Properties.'resource-property'.property | where name -eq 'runtime|memoryCap' | Select '#text'
                            $CPU = $Properties.'resource-property'.property | where name -eq 'hardware|cpuInfo|numCpuPackages' | Select '#text'
                            $CPUcores = $Properties.'resource-property'.property | where name -eq 'hardware|cpuInfo|numCpuCores' | Select '#text'
                            $INFO = $Properties.'resource-property'.property | where name -eq 'cpu|cpuModel' | Select '#text'

                            }

                     }
                    $CheckerOutput = New-Object PsObject -Property @{Name=$vRopsObjName; resourceId=$Result.identifier; resourceKindKey=$Result.resourceKey.resourceKindKey; vCenter=$ParentvCenter.'#text'; Cluster=$ParentCluster.'#text'; Host=$ParentHost.'#text'; State=$PowerState.'#text'; Memory=([Int]$Memory.'#text')/1024/1024; CPU=([Int]$CPU.'#text'); CPUcores=([Int]$CPUcores.'#text'); INFO=$INFO.'#text'}

                    #GetAlerts
                     $ResID = $CheckerOutput.resourceId
                     [xml]$Alerts = $wc.DownloadString("https://$vRopsServer/suite-api/api/alerts?resourceId=$ResID")

                     ForEach ($Alert in $alerts.alerts.alert){

                        $AlertReport += New-Object PSObject -Property @{

                            Name                = $vRopsObjName
                            alertDefinitionName = $Alert.alertDefinitionName
                            alertLevel          = $Alert.alertLevel
                            status              = $Alert.status
                            controlState        = $Alert.controlState
                            startTime           = If ([int64]$Alert.startTimeUTC -gt '') {([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddMilliSeconds([int64]$Alert.startTimeUTC))).tostring("dd/MM/yyyy HH:mm:ss")} else {}
                            cancelTime          = If ([int64]$Alert.cancelTimeUTC -gt '') {([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddMilliSeconds([int64]$Alert.cancelTimeUTC))).tostring("dd/MM/yyyy HH:mm:ss")} else {}
                            updateTime          = If ([int64]$Alert.updateTimeUTC -gt '') {([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddMilliSeconds([int64]$Alert.updateTimeUTC))).tostring("dd/MM/yyyy HH:mm:ss")} else {}
                            suspendUntilTime    = If ([int64]$Alert.suspendUntilTimeUTC -gt ''){([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddMilliSeconds([int64]$Alert.suspendUntilTimeUTC))).tostring("dd/MM/yyyy HH:mm:ss")} else {}
                            alertId             = $Alert.alertId

                        }

                    }

                    Return $CheckerOutput, $AlertReport
                    
                }   
            }
    }  
 }
    else
    {

                     $PropertiesLink = $Checker.resources.resource.links.link | where Name -eq 'latestPropertiesOfResource'
                     $Propertiesurl = 'https://' +$vRopsServer + $PropertiesLink.href
                     [xml]$Properties = $wc.DownloadString($Propertiesurl)

                     switch($Checker.resources.resource.resourceKey.resourceKindKey)
                        {

                        VirtualMachine {

                            $ParentvCenter = $Properties.'resource-property'.property | where name -eq 'summary|parentVcenter' | Select '#text'
                            $ParentCluster = $Properties.'resource-property'.property | where name -eq 'summary|parentCluster' | Select '#text'
                            $ParentHost = $Properties.'resource-property'.property | where name -eq 'summary|parentHost' | Select '#text'
                            $PowerState = $Properties.'resource-property'.property | where name -eq 'summary|runtime|powerState' | Select '#text'
                            $Memory = $Properties.'resource-property'.property | where name -eq 'config|hardware|memoryKB' | Select '#text'
                            $CPU = $Properties.'resource-property'.property | where name -eq 'config|hardware|numCpu' | Select '#text'
                            $INFO = $Properties.'resource-property'.property | where name -eq 'config|guestFullName' | Select '#text'

                            }


                        HostSystem {

                            $ParentvCenter = $Properties.'resource-property'.property | where name -eq 'summary|parentVcenter' | Select '#text'
                            $ParentCluster = $Properties.'resource-property'.property | where name -eq 'summary|parentCluster' | Select '#text'
                            $ParentHost = $Properties.'resource-property'.property | where name -eq 'summary|parentHost' | Select '#text'
                            $PowerState = $Properties.'resource-property'.property | where name -eq 'runtime|powerState' | Select '#text'
                            $Memory = $Properties.'resource-property'.property | where name -eq 'runtime|memoryCap' | Select '#text'
                            $CPU = $Properties.'resource-property'.property | where name -eq 'hardware|cpuInfo|numCpuPackages' | Select '#text'
                            $CPUcores = $Properties.'resource-property'.property | where name -eq 'hardware|cpuInfo|numCpuCores' | Select '#text'
                            $INFO = $Properties.'resource-property'.property | where name -eq 'cpu|cpuModel' | Select '#text'

                            }

                     }
    
    $CheckerOutput = New-Object PsObject -Property @{Name=$vRopsObjName; resourceId=$Checker.resources.resource.identifier; resourceKindKey=$Checker.resources.resource.resourceKey.resourceKindKey; vCenter=$ParentvCenter.'#text'; Cluster=$ParentCluster.'#text'; Host=$ParentHost.'#text'; State=$PowerState.'#text'; Memory=([Int]$Memory.'#text')/1024/1024; CPU=([Int]$CPU.'#text'); CPUcores=([Int]$CPUcores.'#text'); INFO=$INFO.'#text'}

                    #GetAlerts
                     $ResID = $CheckerOutput.resourceId
                     [xml]$Alerts = $wc.DownloadString("https://$vRopsServer/suite-api/api/alerts?resourceId=$ResID")

                     ForEach ($Alert in $alerts.alerts.alert){

                        $Alertreport += New-Object PSObject -Property @{

                            Name                = $vRopsObjName
                            alertDefinitionName = $Alert.alertDefinitionName
                            alertLevel          = $Alert.alertLevel
                            status              = $Alert.status
                            controlState        = $Alert.controlState
                            startTime           = If ([int64]$Alert.startTimeUTC -gt '') {([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddMilliSeconds([int64]$Alert.startTimeUTC))).tostring("dd/MM/yyyy HH:mm:ss")} else {}
                            cancelTime          = If ([int64]$Alert.cancelTimeUTC -gt '') {([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddMilliSeconds([int64]$Alert.cancelTimeUTC))).tostring("dd/MM/yyyy HH:mm:ss")} else {}
                            updateTime          = If ([int64]$Alert.updateTimeUTC -gt '') {([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddMilliSeconds([int64]$Alert.updateTimeUTC))).tostring("dd/MM/yyyy HH:mm:ss")} else {}
                            suspendUntilTime    = If ([int64]$Alert.suspendUntilTimeUTC -gt ''){([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddMilliSeconds([int64]$Alert.suspendUntilTimeUTC))).tostring("dd/MM/yyyy HH:mm:ss")} else {}
                            alertId             = $Alert.alertId

                        }

                    }

                    Return $CheckerOutput, $AlertReport

    }
}


#Logging Function
Function Log([String]$message, [String]$LogType, [String]$LogFile){
    $date = Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    $message = $date + "`t" + $LogType + "`t" + $message
    $message >> $LogFile
}


#Used to execute a job per element to speed up processing.

$scriptBlock = {

param
(
    $Resource,
    $resourceIdLookupTable,
    $resourceKindKeyLookupTable,
    $elementsLookupTable,
    $rollUpType,
    $intervalType,
    $StartDate,
    $EndDate,
    $Output 
)

#Vars
$StartDateFile = $StartDate.tostring("yyyyMMdd-HHmmss")            
$EndDateFile = $EndDate.tostring("yyyyMMdd-HHmmss")
$MetricOutput = $Output + 'Collected_Metrics_' + $Resource.'resourceId' + '_' + $intervalType + '_' +[String]$StartDateFile + '_' + [String]$EndDateFile + '.csv'

$report = @()

#Slow part of the code... need to make it faster
#----------------------------------------------

  foreach ($node in $Resource.'stat-list'.stat)
   {
   #Collection Date, not run time
   $MetricName = $node.statKey.Key
    $intervalType = $node.intervalUnit.intervalType
    $rollUpType = $node.rollUpType
    $Values     = @($node.data -split ' ')
    $Timestamps = @($node.timestamps -split ' ')

        for ($i=0; $i -lt $Values.Count -and $i -lt $Timestamps.Count; $i++) {
            $report += New-Object PSObject -Property @{
            METRIC     = $MetricName
            resourceId = $Resource.'resourceId'
            Timestamp  = ([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddMilliSeconds([int64]$Timestamps[$i]))).tostring("dd/MM/yyyy HH:mm:ss")
            intervalType = $intervalType
            rollUpType = $rollUpType
            value      = $Values[$i]
            }
        }

    }

#----------------------------------------------

#Add $resourceId, $resourceKindKey & Friendly Name
$report  | Sort-Object -Property resourceId | ForEach-Object {
    $_ | Add-Member -MemberType NoteProperty -Name resourceName -Value $resourceIdLookupTable."$($_.resourceId)" -PassThru
} | Export-csv $MetricOutput -NoTypeInformation

}

switch($CollectionType)
    {

Collection {





Log -Message "Collecting $intervalType between $StartDateFile and $EndDateFile, running $maxJobCount job's at a time" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
Log -Message "Data collection for object: $ResourceByName" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc

if ($Metrics -gt ""){


Log -Message "Collecting the Metrics: $Metrics" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc

}
else {

#$Metrics = "cpu|usagemhz_average"

#Log -Message "No metrics were specified, defaulting to cpu|usagemhz_average" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
#echo "No metrics were specified, defaulting to cpu|usagemhz_average"

}

#Take all certs.
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

[int64]$StartDateEpoc = Get-Date -Date $StartDate.ToUniversalTime() -UFormat %s
$StartDateEpoc = $StartDateEpoc*1000 
[int64]$EndDateEpoc = Get-Date -Date $EndDate.ToUniversalTime() -UFormat %s
$EndDateEpoc = $EndDateEpoc*1000 

#Lookup Name and map to resourceId Table

$AlarmTable = @()
$ObjectLookupTable = @()

Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
echo 'Looking stuff up'
Log -Message "Looking stuff up" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc

ForEach ($ResourceName in $ResourceByName){


    #Map name to resourceId for lookup
    $resourceLookup = GetObject $ResourceName $vRopsAddress $vRopsUser $vRopsPassword

            If ($resourceLookup[0].resourceId -gt ''){

                #Generate String for resourceId URL lookup.
                $ResourceByNameSearcString += 'resourceId='+$resourceLookup[0].resourceId+'&'
                $ResourceParentvCenter = $resourceLookup[0].vCenter
                $ResourceParentCluster = $resourceLookup[0].Cluster
                $ResourceParentHost = $resourceLookup[0].Host
                $ResourcePowerState = $resourceLookup[0].State
                $ResourceMemory = $resourceLookup[0].Memory
                $ResourceCPU = $resourceLookup[0].CPU
		        $ResourceCPUcores = $resourceLookup[0].CPUcores
                $ResourceINFO = $resourceLookup[0].INFO

                ForEach ($alarm in $resourceLookup[1]){

                    $AlarmTable += New-Object PSObject -Property @{
                    name                = $ResourceName
                    controlState		= $alarm.controlState
                    suspendUntilTime	= $alarm.suspendUntilTime
                    cancelTime		    = $alarm.cancelTime
                    updateTime		    = $alarm.updateTime
                    alertLevel		    = $alarm.alertLevel
                    alertId			    = $alarm.alertId
                    alertDefinitionName	= $alarm.alertDefinitionName
                    startTime		    = $alarm.startTime
                    status			    = $alarm.status
                    }
                 }
                }

            else {
                
                Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
                echo "I Still cant find $ResourceName, are you looking in the correct environment? does the machine still exist?"
                Log -Message "I Still cant find $ResourceName, does the machine exist?" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
                
                }

    $ObjectLookupTable += New-Object PSObject -Property @{
        resourceId = $resourceLookup[0].resourceId
        resourceName = $ResourceName
        resourceKindKey = $resourceLookup[0].resourceKindKey
        ParentvCenter = $ResourceParentvCenter
        ParentCluster = $ResourceParentCluster
        ParentHost = $ResourceParentHost
        PowerState = $ResourcePowerState
        Memory = $ResourceMemory
        CPU = $ResourceCPU
	CPUcores = $ResourceCPUcores
        INFO = $ResourceINFO
        }
}

Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
echo 'Finished mapping UUID to ResourceID'
Log -Message "Finished mapping UUID to ResourceID" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc

If ($ResourceByNameSearcString  -gt ''){

if ($Metrics -gt ''){

    ForEach ($Metric in $Metrics){

    $MetricLookupString += '&statKey=' + $Metric

    }

    $url = 'https://'+$vRopsAddress+'/suite-api/api/resources/stats?'+ $ResourceByNameSearcString + 'rollUpType='+ $rollUpType + '&intervalType=' + $intervalType + $MetricLookupString + '&begin=' + $StartDateEpoc + '&end=' + $EndDateEpoc
 
}

else {

    $url = 'https://'+$vRopsAddress+'/suite-api/api/resources/stats?'+ $ResourceByNameSearcString + 'rollUpType='+ $rollUpType + '&intervalType=' + $intervalType + '&begin=' + $StartDateEpoc + '&end=' + $EndDateEpoc
}

$webcall = new-object system.net.WebClient
$webcall.Credentials = new-object System.Net.NetworkCredential($vRopsUser, $vRopsPassword)

}

else 

{

Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
echo 'No Machines to look up, terminating script'
Log -Message "No Machines to look up, terminating script" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
Exit
}

Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
echo 'Call URL to Download the XML'
Log -Message "Call URL to Download the XML" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc

#Download data in XML from vRops

[xml]$Data = $webcall.DownloadString($url)

Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
echo 'XML data now stored in Variable'
Log -Message "XML data now stored in Variable" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc

  if($Format -eq 'XML'){
    #Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    echo 'Data Downloaded to XML file'
    Log -Message "Data Downloaded to XML file" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    $OutputXMLFile = $ScriptPath + '\completed\Collected_Metrics_' + $intervalType + '_' +[String]$StartDateFile + '_' + [String]$EndDateFile + '_' + $RunDateTime + '.xml'
    $ShareXMLlFile = $share + 'Collected_Metrics_' + $intervalType + '_' +[String]$StartDateFile + '_' + [String]$EndDateFile + '_' + $RunDateTime + '.xml'
    $webcall.DownloadFile($url, $OutputXMLFile)
    Log -Message "Task complete, pickup your file from the completed folder" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    Echo "Task complete, pickup your file from the completed folder"
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    echo 'All done, Terminating Script'
    Log -Message "All done, Terminating Script" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    #Return $OutputXMLFile
    Remove-Variable * -ErrorAction SilentlyContinue
    remove-item $Output -Force
    Exit
  }


Log -Message "Running $url" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
#echo $url

#Add Name to resourceID Mapping
$ObjectLookupTable | Sort-Object -Property resourceId | ForEach-Object -Begin {
    $resourceIdLookupTable = @{}
} -Process {
    $resourceIdLookupTable.Add($_.resourceId,$_.resourceName)
}

$jobQueue = New-Object System.Collections.ArrayList

$resources = $Data.'stats-of-resources'
$UUIDS = $Resource.'resourceId'

Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
echo 'Processing the XML into a CSVs'
Log -Message "Processing the XML into a CSVs" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc

        # Create our job queue.


            # Main loop of the script.  
            # Loop through each VM and start a new job if we have less than $maxJobCount outstanding jobs.  
            # If the $maxJobCount has been reached, sleep 3 seconds and check again. 


Foreach ($Resource in $Resources.'stats-of-resource'){

              # Wait until job queue has a slot available.
              while ($jobQueue.count -ge $maxJobCount) {
                echo "jobQueue count is $($jobQueue.count): Waiting for jobs to finish before adding more."
                foreach ($jobObject in $jobQueue.toArray()) {
            	    if ($jobObject.job.state -eq 'Completed') { 
            	      echo "jobQueue count is $($jobQueue.count): Removing job"
            	      $jobQueue.remove($jobObject) 		
            	    }
            	  }
            	sleep $sleepTimer
              }  
  
              echo "jobQueue count is $($jobQueue.count): Adding new job: $($Resource.'resourceId')"
              

              $job = Start-Job -name $Resource.'resourceId' -ScriptBlock $scriptBlock -ArgumentList $Resource, $resourceIdLookupTable, $resourceKindKeyLookupTable, $elementsLookupTable, $rollUpType, $intervalType, $StartDate, $EndDate, $Output            

              $jobObject          = "" | select Element, job
              $jobObject.Element  = $Element
              $jobObject.job      = $job
              $jobQueue.add($jobObject) | Out-Null
            }

Get-Job | Wait-Job | Out-Null

Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
echo 'Finished Generating the CSVs'
Log -Message "Finished Generating the CSVs" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc


Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
echo 'Start merge of the CSVs into a variable'
Log -Message "Start merge of the CSVs into a variable" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc

#get all CSV files from this job

$Filearr = Get-ChildItem $Output | 
       Where-Object {$_.Name -like '*.csv'} | 
       Foreach-Object { $Output + $_.Name}


#Merge all CSV data from this job


$OutputMerge = @();            
foreach($CSV in $filearr) {            
    if(Test-Path $CSV) {            
                    
        $FileName = [System.IO.Path]::GetFileName($CSV)            
        $temp = Import-CSV -Path $CSV | select *      
        $OutputMerge += $temp            
            
    } else {            
        Write-Warning "$CSV : No such file found"            
    }            
            
}            



#Delete individual CSVs

remove-item $Filearr -Force

#Output merge to Excel
  if($Format -eq 'CSV'){
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    echo 'CSV Merge complete, deleting individual CSVs and saving to single CSV'
    Log -Message "CSV Merge complete, deleting individual CSVs and saving to single CSV" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    $OutputCSVFile = $ScriptPath + '\completed\Collected_Metrics_' + $intervalType + '_' +[String]$StartDateFile + '_' + [String]$EndDateFile + '_' + $RunDateTime + '.csv'
    $ShareCSVlFile = $share + 'Collected_Metrics_' + $intervalType + '_' +[String]$StartDateFile + '_' + [String]$EndDateFile + '_' + $RunDateTime + '.csv'
    $OutputMerge | Sort-Object { $_.Timestamp -as [datetime] } | export-csv $OutputCSVFile -NoTypeInformation
    Log -Message "Task complete, pickup your file from the completed folder" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    Echo "Task complete, pickup your file from the completed folder"
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    echo 'All done, Terminating Script'
    Log -Message "All done, Terminating Script" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    remove-item $Output -Force
    Remove-Variable * -ErrorAction SilentlyContinue
    Exit
  }


#Output merge to Excel
  if($Format -eq 'XLS'){
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    echo 'CSV Merge complete, deleting individual CSVs and creating Excel file'
    Log -Message "CSV Merge complete, deleting individual CSVs and creating Excel file" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    $OutputExcelFile = $ScriptPath + '\completed\Collected_Metrics_' + $intervalType + '_' +[String]$StartDateFile + '_' + [String]$EndDateFile + '_' + $RunDateTime + '.xlsx'
    $ShareExcelFile = $share + 'Collected_Metrics_' + $intervalType + '_' +[String]$StartDateFile + '_' + [String]$EndDateFile + '_' + $RunDateTime + '.xlsx'
    $OutputMerge | select Timestamp,value,METRIC,resourceName,HOSTNAME | Sort-Object { $_.Timestamp -as [datetime] } | Export-Excel $OutputExcelFile -WorkSheetname Data -ChartType Line -IncludePivotChart -IncludePivotTable -PivotRows Timestamp -PivotData value -PivotColumns resourceName,METRIC
    $ObjectLookupTable | select resourceName,resourceKindKey,ParentvCenter,ParentCluster,ParentHost,PowerState,Memory,CPU,CPUcores,INFO | Export-Excel $OutputExcelFile -WorkSheetname Config
    $AlarmTable | Sort-Object { $_.startTime -as [datetime] } | select Name,startTime,updateTimealertLevel,suspendUntilTime,cancelTime,alertId,alertDefinitionName,status,controlState | Export-Excel $OutputExcelFile -WorkSheetname Alarms
    Log -Message "Task complete, pickup your file from the completed folder" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    Echo "Task complete, pickup your file from the completed folder"
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    echo 'All done, Terminating Script'
    Log -Message "All done, Terminating Script" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    remove-item $Output -Force
    Remove-Variable * -ErrorAction SilentlyContinue
    Exit
  }

remove-item $Output -Force
Remove-Variable * -ErrorAction SilentlyContinue

}

default{"Usage

The script can be run by specifying all parameters otherwise it will use some default values.

-vRopsAddress : IP or DNS name of vRops environment to pull data from.

-Creds : HOME is the name of the XML file to pull the credentials from IE: HOME.xml

-CollectionType : Collection (this will be used to switch between different types of collections in the future).

-ResourceByName : is an array of objects (datastores, hosts, VM's etc...) use the name of the object as it appears in vRops / vCenter.

-Metrics : is an array of specific metrics to collect from vRops for the object(s), vRops has 100's of metrics of each object and they can be instance specific so the best way to collect a list is to run a daily collection for 1 day without specifing -Metric and filtering the results in Excel. (Example below)
.\DumpvRopsMetric.ps1 -vRopsAddress vRops.vMan.ch -CollectionType Collection -ResourceByName 'WINSRV2','WINSRV4' -Creds HOME -rollUpType AVG -intervalType DAYS -startdate '2016/06/21' -enddate '2016/06/21'
 
-Interval and -rollUpTypes must be specified together.
intervalType=HOURS (rollUpType=SUM,AVG,MIN,MAX,LATEST,COUNT)
intervalType=MINUTES (rollUpType=SUM,AVG,MIN,MAX,LATEST,COUNT)
intervalType=SECONDS 
intervalType=DAYS (rollUpType=SUM,AVG,MIN,MAX,LATEST,COUNT)
intervalType=WEEKS (rollUpType=SUM,AVG,MIN,MAX,LATEST,COUNT)
intervalType=MONTHS (rollUpType=SUM,AVG,MIN,MAX,LATEST,COUNT)
intervalType=YEARS (rollUpType=SUM,AVG,MIN,MAX,LATEST,COUNT)

Run with -Metric, -startdate and -enddate
.\DumpvRopsMetric.ps1 -vRopsAddress vRops.vMan.ch -CollectionType Collection -ResourceByName 'WINSRV2','WINSRV4' -Metrics 'cpu|usagemhz_average','cpu|costopPct','cpu|readyPct','cpu|iowaitPct','cpu|idletimepercent','cpu|demandPct' -Creds HOME -rollUpType AVG -intervalType MINUTES -startdate '2016/09/18 19:20' -enddate '2016/09/19 19:20' -Format XLS

Run without -startdate & -enddate will default to the last 24H
.\DumpvRopsMetric.ps1 -vRopsAddress vRops.vMan.ch -CollectionType Collection -ResourceByName 'WINSRV2','WINSRV4' -Metrics 'cpu|usagemhz_average','mem|usage_average','cpu|perCpuCoStopPct' -Creds HOME -rollUpType AVG -intervalType HOURS -Format XLS

Run without -Metrics and it will pull all metrics and state data
.\DumpvRopsMetric.ps1 -vRopsAddress vRops.vMan.ch -CollectionType Collection -ResourceByName 'WINSRV2','WINSRV4' -Creds HOME -rollUpType AVG -intervalType DAYS -startdate '2016/08/04 08:00' -enddate '2016/08/05 10:00' -Format XLS
        "}
}