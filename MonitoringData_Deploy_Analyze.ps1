#******************************************************************** LOAD PARAMETERS FROM Notifications_AreaResume_GUI.PS1 ***************************************************************************
Param
 (
	[String]$Selected_item_Name,
	[String]$Monitoring_Host_Content	
	
 )
 
																															
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  				| out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 				| out-null

$Deployment_ComputerName = $Selected_item_Name
																															
$Global:Current_Folder =(get-location).path 
[System.Windows.Forms.MessageBox]::Show("$Monitoring_Host_Content")

										
#************************************************************************** GET MONITORING INFOS FROM XML ******************************************************************************************************		
# If ($Mail -eq $true)
	# {
		# $Progdata = $env:PROGRAMDATA
		# $Global:List_XML_Content = "$Progdata\Monitoring_Infos.xml"						
		# $Input_XML = [xml] (Get-Content $List_XML_Content)	 
		# foreach ($infos in $Input_XML.selectNodes("Mail_Infos"))		
			# {
				# $Global:Mail_SMTP = $infos.SMTP			
				# $Global:Mail_From = $infos.MailFrom			
				# $Global:Mail_To = $infos.MailTo			
			# }		
	# }
		
	
$Global:URL = "http://" + "$Monitoring_Host_Content" + ":9801/MDTMonitorData/Computers/"					

Write-Host "##############################################################################" -ForegroundColor Cyan 
Write-Host "Deployment statut analyzer" -ForegroundColor yellow 
Write-Host "The script is analyzing the deployment of the following computer: $Deployment_ComputerName" -ForegroundColor yellow 
Write-Host "Do not close this script !!!" -ForegroundColor yellow 
Write-Host "Once deployment is finished a GUI will be displayed" -ForegroundColor yellow 
Write-Host "##############################################################################" -ForegroundColor Cyan 
	
	

function GetMDTData { 
  $Data = Invoke-RestMethod $URL
  foreach($property in ($Data.content.properties)) 
  { 
		$Percent = $property.PercentComplete.'#text' 		
		$Current_Steps = $property.CurrentStep.'#text'			
		$Total_Steps = $property.TotalSteps.'#text'		
		
		$Step_Name = $property.StepName		
		If ($Percent -eq 100)
			{
				$Global:StepName = "Deployment finished"
				$Percent_Value = $Percent + "%"				
			}
		Else
			{
				If ($Step_Name -eq "")
					{					
						If ($Percent -gt 0) 					
							{
								$Global:StepName = "Computer restarted"
								$Percent_Value = $Percent + "%"
							}	
						Else							
							{
								$Global:StepName = "Deployment not started"	
								$Percent_Value = "Not started"	
							}
					}
				Else
					{
						$Global:StepName = $property.StepName		
						$Percent_Value = $Percent + "%"					
					}					
			}

		$Deploy_Status = $property.DeploymentStatus.'#text'					
		If (($Percent -eq 100) -and ($Step_Name -eq "") -and ($Deploy_Status -eq 1))
			{
				$Global:StepName = "Running in PE"						
			}			
			
			
		$End_Time = $property.EndTime.'#text' 	
		If ($End_Time -eq $null)
			{
				If ($Percent -eq $null)
					{									
						$EndTime = "Not started"
						$Ellapsed = "Not started"												
					}
				Else
					{
						$EndTime = "Not finished"
						$Ellapsed = "Not finished"					
					}
			}
		Else
			{
				$EndTime = ([datetime]$($property.EndTime.'#text')).ToLocalTime().ToString('HH:mm:ss')  	 
				$Ellapsed = new-timespan -start ([datetime]$($property.starttime.'#text')).ToString('HH:mm:ss') -end ([datetime]$($property.endTime.'#text')).ToString('HH:mm:ss'); 				
			}
	
    New-Object PSObject -Property @{ 
      ComputerName = $($property.Name); 
      Percent_Complete = $Percent_Value; 	  
      Step_Name = $StepName;	  	  
      Actual_Step = $property.CurrentStep.'#text'	 	  
      All_my_Steps = $property.TotalSteps.'#text'		  
      Warnings = $($property.Warnings.'#text'); 	  
      Errors = $($property.Errors.'#text'); 	  
      ID = $($property.ID.'#text'); 
      LastTime = $($property.LastTime.'#text'); 
      DeploymentStatus = $($property.DeploymentStatus.'#text'); 	  
      "Deployment Status" = $( 
        Switch ($property.DeploymentStatus.'#text') { 
        1 { "Running" } 
        2 { "Failed" } 
        3 { "Success" } 
        4 { "Unresponsive" } 		
        Default { "Unknown" } 
        } 
      ); 	  
      "Date" = $($property.StartTime.'#text').split("T")[0]; 
      Start_time = ([datetime]$($property.StartTime.'#text')).ToLocalTime().ToString('HH:mm:ss')  
	  End_time = $EndTime;
      Ellapsed_time = $Ellapsed;	  	  
    } 
  } 
}

$Global:MyData = GetMDTData | Select Date, ComputerName, Actual_Step, All_my_Steps, LastTime, Percent_Complete, Step_Name, Warnings, Errors, Start_time, End_time, Ellapsed_time, DeploymentStatus, "Deployment Status" | where {$_.ComputerName -eq $Deployment_ComputerName}	

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") 
[reflection.assembly]::loadwithpartialname("System.Drawing")
$path = Get-Process -id $pid | Select-Object -ExpandProperty Path            		
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)    		
$notify = new-object system.windows.forms.notifyicon
$notify.icon = $icon
$notify.visible = $true		

Do
	{			
		$Deployment_Status = $MyData.DeploymentStatus		
		$Deployment_Date = $MyData.Date
		$Deployment_CompName = $MyData.ComputerName
		$Deployment_LastTime = $MyData.LastTime
		$Deployment_StepName = $MyData.Step_Name
		$Deployment_Warnings = $MyData.Warnings
		$Deployment_Errors = $MyData.Errors
		$Deployment_Percent_Complete = $MyData.Percent_Complete		
		$Deployment_Start_time = $MyData.Start_time		
		$Deployment_End_time = $MyData.End_time
		$Deployment_Ellapsedtime = $MyData.Ellapsed_time

		 If ($Deployment_Status -eq 4)		 
			{
				$message = "Deployment is unresponsive on $Deployment_ComputerName.`nPlease check status of the computer."
				$notify.showballoontip(10,$title,$Message, `
				[system.windows.forms.tooltipicon]::error)						
				break		
			}

		ElseIf ($Deployment_Status -eq 3)		
			{
				$message = "Deployment success on $Deployment_ComputerName"
				$notify.showballoontip(10,$title,$Message, `
				[system.windows.forms.tooltipicon]::info)												
				break
			}	
			
		ElseIf ($Deployment_Status -eq 2)		
			{		
				$message = "Deployment is unresponsive on $Deployment_ComputerName.`nPlease check status of the computer."
				$notify.showballoontip(10,$title,$Message, `
				[system.windows.forms.tooltipicon]::error)								
				break
			}				
	} 
	
While ($MyData.Percent_Complete	 -lt 100) 
write-host "Deployment finished"
	

	
	
	


