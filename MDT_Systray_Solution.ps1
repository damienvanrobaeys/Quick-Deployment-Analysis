Param
 (
	[String]$Restart	
 )

If ($Restart -ne "") 
	{
		sleep 10
	}
 
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 		 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | out-null

$Current_Folder = split-path $MyInvocation.MyCommand.Path

$Global:Current_Folder = $PSScriptRoot
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\mmc.exe")
$Current_Version = "1.1"

################################################################################################################################"
# DIFFERENT GUI DISPLAY PART
################################################################################################################################"

# ----------------------------------------------------
# Part - Main Monitoring list
# ----------------------------------------------------

[xml]$xaml =  
@"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
WindowStyle="None" 
Height="600" 
Width="400"
ResizeMode="NoResize" 
ShowInTaskbar="False"
AllowsTransparency="True" 
Background="Transparent"
>
<Border  BorderBrush="Black" BorderThickness="1" Margin="10,10,10,10">

<Grid Name="grid" Background="White">
	<StackPanel Margin="0,0,0,0" >

		<StackPanel VerticalAlignment="Center">
			<Image Width="182" Height="70" Source="$Current_Folder\images\logo.png" ></Image>	
		</StackPanel>	
		
		<StackPanel HorizontalAlignment="Center" Margin="0,0,0,0">
			<Label Content="Current installations" Foreground="Black" FontWeight="Bold" FontSize="18"/>			
		</StackPanel>	
		
		<StackPanel Orientation="Horizontal" HorizontalAlignment="Center">		
				<Label x:Name="Success_Comp" Foreground="Black" FontSize="13"/>				
				<Label x:Name="Window_Running_Comp" Foreground="Black" FontSize="13"/>
				<Label x:Name="Failed_Comp" Foreground="Black" FontSize="13"/>					
				<Label x:Name="Unresponsive_Comp" Foreground="Black" FontSize="13"/>		
		</StackPanel>		

		<StackPanel Height="407" VerticalAlignment="Center" HorizontalAlignment="Center">		
			<StackPanel Margin="0,20,0,0" Orientation="Vertical" VerticalAlignment="Center">

				<StackPanel Margin="0,0,0,0">	
					<StackPanel x:Name="DataGrid_Status_On">						
						<DataGrid RowHeaderWidth="0" SelectionMode="Single"  Background="#313130" Height="350"   AutoGenerateColumns="False" Name="DataGrid_Monitoring"  ItemsSource="{Binding}"  Margin="2,2,2,2" >
	 
						<DataGrid.Resources>
							<Style BasedOn="{StaticResource {x:Type DataGridColumnHeader}}" TargetType="{x:Type DataGridColumnHeader}">
								<Setter Property="Background" Value="Transparent" />
								<Setter Property="FontSize" Value="14" />
								<Setter Property="Foreground" Value="White" />
								<Setter Property="BorderBrush" Value="Red" />

							</Style>
						</DataGrid.Resources>				

						<DataGrid.RowStyle>
							<Style TargetType="DataGridRow"> 
								<Setter Property="IsHitTestVisible" Value="False"/>						
								<Style.Triggers>
									<DataTrigger Binding="{Binding DeploymentStatus}" Value="Running">
										<Setter Property="Background" Value="#5290E9"></Setter>
									</DataTrigger>			
									<DataTrigger Binding="{Binding DeploymentStatus}" Value="Failed">
										<Setter Property="Background" Value="#E14D57"></Setter>
									</DataTrigger>
									<DataTrigger Binding="{Binding DeploymentStatus}" Value="Unresponsive">
										<Setter Property="Background" Value="#EC932F"></Setter>
									</DataTrigger>				
									<DataTrigger Binding="{Binding DeploymentStatus}" Value="Success">
										<Setter Property="Background" Value="#71B37C"></Setter>
									</DataTrigger>
								</Style.Triggers>
							</Style>
						</DataGrid.RowStyle>	

						<DataGrid.Columns>	
								<DataGridTextColumn Width="auto" Header="Computer Name" Binding="{Binding Name}"/>							
								<DataGridTextColumn Width="auto" Header="Status" Binding="{Binding DeploymentStatus}"/>							
								<DataGridTextColumn Width="auto" Header="Percent" Binding="{Binding PercentComplete}"/>
							</DataGrid.Columns>
						</DataGrid>		
					</StackPanel>	


					<StackPanel x:Name="DataGrid_Status_Off">		
						<Image Margin="0,60,0,0" Width="150" Height="150" Source="$Current_Folder\images\KO.png" ></Image>							
					</StackPanel>				
				</StackPanel>
			</StackPanel>		
		</StackPanel>

		<StackPanel Name="Main_Monitoring_Status" Height="40" Background="#F2F2F2" Width="400" HorizontalAlignment="Center" VerticalAlignment="Bottom">	
				<Line Grid.Row="1" X1="0" Y1="0" X2="1"  Y2="0" Stroke="Black" StrokeThickness="0.2" Stretch="Uniform"></Line>						
				<Label FontWeight="Bold" Foreground="White" Name="Main_Monitoring_msg" Margin="0,7,0,0"  FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center"/>
		</StackPanel>	

	</StackPanel>		
</Grid>	
</Border>
</Window>
"@

$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$DataGrid_Status_On = $window.findname("DataGrid_Status_On") 
$DataGrid_Status_Off = $window.findname("DataGrid_Status_Off") 
$DataGrid_Monitoring = $window.findname("DataGrid_Monitoring") 
$Success_Comp = $window.findname("Success_Comp") 
$Window_Running_Comp = $window.findname("Window_Running_Comp") 
$Failed_Comp = $window.findname("Failed_Comp") 
$Unresponsive_Comp = $window.findname("Unresponsive_Comp") 
$Main_Monitoring_Status = $window.findname("Main_Monitoring_Status") 
$Main_Monitoring_msg = $window.findname("Main_Monitoring_msg") 



# ----------------------------------------------------
# Part - Monitoring host settings
# ----------------------------------------------------

[xml]$xaml_config =  
@"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
WindowStyle="None" 
Height="260" 
Width="380"
ResizeMode="NoResize" 
ShowInTaskbar="False"
Background="Transparent"
AllowsTransparency="True" 
>

<Border  BorderBrush="Black" BorderThickness="1" Margin="10,10,10,10">
<Grid Name="grid" Background="White" >		
	<StackPanel HorizontalAlignment="Center">
		<StackPanel Margin="0,0,0,0" >
			<StackPanel Margin="0,0,0,0" Height="200" Orientation="Vertical">		
				<StackPanel Margin="0,0,0,0"  Orientation="Horizontal">					
					<Image Margin="10,0,0,0" Width="30" Height="30" Source="$Current_Folder\images\settings2.png" ></Image>				
					<Label Margin="5,0,0,0" Content="Monitoring Host Settings" FontWeight="Bold" Foreground="Black" FontSize="14"/>	
				</StackPanel>		
				
				<StackPanel Margin="0,0,0,0"  HorizontalAlignment="Center">	
					<Line Grid.Row="0" X1="0" Y1="0" X2="1"  Y2="0" Stroke="White" StrokeThickness="0.5" Stretch="Uniform" />									
					<Label Margin="0,15,0,0" Content="Type your Monitoring Host" Foreground="Black" FontSize="18" HorizontalAlignment="Center"/>
					<TextBox Width="250" AcceptsReturn="True" TextWrapping="Wrap" Name="Host_TxtBox" Height="25" FontSize="16" />			
					<Button Width="250" Name="Set_Host" Height="25" Content="Set this host" FontSize="16" Margin="0,5,0,0"/>		
				</StackPanel>						
			</StackPanel>
			
			
			<StackPanel Name="MonitoringHost_Status_Block" Height="40" Background="#F2F2F2" Width="370" HorizontalAlignment="Center" VerticalAlignment="Bottom">	
					<Line Grid.Row="1" X1="0" Y1="0" X2="1"  Y2="0" Stroke="Black" StrokeThickness="0.2" Stretch="Uniform"></Line>						
					<Label FontWeight="Bold" Foreground="White" Name="MonitoringHost_Status_Label" Margin="0,10,0,0"  FontSize="10" HorizontalAlignment="Center" VerticalAlignment="Center"/>
			</StackPanel>				
		</StackPanel>		
	</StackPanel>		
</Grid>	
</Border>
</Window>
"@

$window_Config = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml_config))
$Host_TxtBox = $window_Config.findname("Host_TxtBox") 
$Set_Host = $window_Config.findname("Set_Host") 
$MonitoringHost_Status_Label = $window_Config.findname("MonitoringHost_Status_Label") 
$MonitoringHost_Status_Block = $window_Config.findname("MonitoringHost_Status_Block") 




# ----------------------------------------------------
# Part - Be notified when computers are installed
# ----------------------------------------------------

[xml]$xaml_Running =  
@"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
WindowStyle="None" 
Height="600" 
Width="400"
ResizeMode="NoResize" 
ShowInTaskbar="False"
AllowsTransparency="True" 
Background="Transparent"
>
<Border  BorderBrush="Black" BorderThickness="1" Margin="10,10,10,10">

<Grid Name="grid" Background="White">
	<StackPanel Margin="0,10,0,0" >
		<StackPanel VerticalAlignment="Center">
			<Image Width="112" Height="100" Source="$Current_Folder\images\logo.png" ></Image>	
		</StackPanel>	
		
		<StackPanel HorizontalAlignment="Center" Margin="0,10,0,0">
			<Label Content="Running installations" Foreground="Black" FontWeight="Bold" FontSize="18"/>			
		</StackPanel>	

		<StackPanel Height="414" VerticalAlignment="Center" HorizontalAlignment="Center">		
			<StackPanel Margin="0,20,0,0" Orientation="Vertical" VerticalAlignment="Center">
				<StackPanel Margin="0,0,0,0">	
					<StackPanel x:Name="DataGrid_Running_On">						
						<DataGrid  Width="auto" RowHeaderWidth="0" SelectionMode="Single"  Background="#313130" Height="350"   AutoGenerateColumns="True" Name="DataGrid_Running"  ItemsSource="{Binding}"  Margin="2,2,2,2" >
	 
						<DataGrid.Resources>
							<Style BasedOn="{StaticResource {x:Type DataGridColumnHeader}}" TargetType="{x:Type DataGridColumnHeader}">
								<Setter Property="Background" Value="Transparent" />
								<Setter Property="FontSize" Value="14" />
								<Setter Property="Foreground" Value="White" />
								<Setter Property="BorderBrush" Value="Red" />
							</Style>
						</DataGrid.Resources>				

						<DataGrid.Columns>	
								<DataGridTextColumn Width="auto" Header="Computer Name" Binding="{Binding Name}"/>							
								<DataGridTextColumn Width="auto" Header="Status" Binding="{Binding DeploymentStatus}"/>							
								<DataGridTextColumn Width="auto" Header="Percent" Binding="{Binding PercentComplete}"/>
							</DataGrid.Columns>
						</DataGrid>		
						
						<Button Content="Notify me for the selected computer" Name="NotifyMe" Height="25" Margin="0,5,0,0"/>
					</StackPanel>	

					<StackPanel x:Name="DataGrid_Running_Off">		
						<Image Margin="0,60,0,0" Width="150" Height="150" Source="$Current_Folder\images\KO.png" ></Image>							
					</StackPanel>				
				</StackPanel>				
			</StackPanel>		
		</StackPanel>	
		
		<StackPanel Name="Main_Running_Monitoring_Status" Height="40" Background="#F2F2F2" Width="400" HorizontalAlignment="Center" VerticalAlignment="Bottom">	
				<Line Grid.Row="1" X1="0" Y1="0" X2="1"  Y2="0" Stroke="Black" StrokeThickness="0.2" Stretch="Uniform"></Line>						
				<Label FontWeight="Bold" Foreground="White" Name="Main_Running_Monitoring_msg" Margin="0,7,0,0"  FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center"/>
		</StackPanel>			
		
		
	</StackPanel>	
</Grid>	
</Border>
</Window>
"@

$Window_Part_Running_Comp = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml_Running))
$DataGrid_Running = $Window_Part_Running_Comp.findname("DataGrid_Running") 
$DataGrid_Running_On = $Window_Part_Running_Comp.findname("DataGrid_Running_On") 
$DataGrid_Running_Off = $Window_Part_Running_Comp.findname("DataGrid_Running_Off") 
$Main_Running_Monitoring_Status = $Window_Part_Running_Comp.findname("Main_Running_Monitoring_Status") 
$Main_Running_Monitoring_msg = $Window_Part_Running_Comp.findname("Main_Running_Monitoring_msg") 
$NotifyMe = $Window_Part_Running_Comp.findname("NotifyMe") 




# ----------------------------------------------------
# Part - About
# ----------------------------------------------------

[xml]$xaml_About =  
@"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
WindowStyle="None" 
Height="260" 
Width="380"
ResizeMode="NoResize" 
ShowInTaskbar="False"
Background="Transparent"
AllowsTransparency="True" 
>

<Border  BorderBrush="Black" BorderThickness="1" Margin="10,10,10,10">
<Grid Name="grid" Background="White" >		
	<StackPanel HorizontalAlignment="Center"  VerticalAlignment="Center" Orientation="Vertical">
		<StackPanel Height="200"  HorizontalAlignment="Center"  VerticalAlignment="Top">		
			<Image Width="120" Height="100" Source="$Current_Folder\images\logoQDA.png" Margin="0,10,0,0"/>					
			<Label Margin="0,10,0,0" x:Name="About_Version" Content="Version: 1.0" Foreground="Black" FontSize="12" HorizontalAlignment="Center"/>
			<Label Margin="0,-5,0,0" x:Name="About_Date" Content="Last release: 18/09/19" Foreground="Black" FontSize="12" HorizontalAlignment="Center"/>				
			<Label Margin="0,-5,0,0" x:Name="About_Name" Content="Damirn van robaeys" Foreground="Black" FontSize="12"  HorizontalAlignment="Center"/>
		</StackPanel>
		
		<StackPanel Height="40" Name="Main_Update_Status" Background="#F2F2F2" Width="370" HorizontalAlignment="Center" VerticalAlignment="Bottom">	
				<Line Grid.Row="1" X1="0" Y1="0" X2="1"  Y2="0" Stroke="Black" StrokeThickness="0.2" Stretch="Uniform"></Line>						
				<Label Margin="0,10,0,0" Name="Update_Status_Label" Content="The tool is up to date" FontSize="10" HorizontalAlignment="Center" VerticalAlignment="Center"/>
		</StackPanel>		
	</StackPanel>		
</Grid>	
</Border>
</Window>
"@

$window_About = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml_About))
$About_Version = $window_About.findname("About_Version") 
$About_Date = $window_About.findname("About_Date") 
$About_Name = $window_About.findname("About_Name") 
$Main_Update_Status = $window_About.findname("Main_Update_Status") 
$Update_Status_Label = $window_About.findname("Update_Status_Label") 

$About_Version.Content = "Version: 1.1"
$About_Date.Content = "Release date: 08/07/2018"
$About_Name.Content = "Author: Damien Van Robaeys"

################################################################################################################################"
# DIFFERENT GUI DISPLAY PART
################################################################################################################################"

# ----------------------------------------------------
# Part - Main Monitoring list
# ----------------------------------------------------

$Progdata = $env:PROGRAMDATA
$Monitoring_Host_Txt = "$Progdata\Monitoring_Host.txt"

If (!(test-path $Monitoring_Host_Txt))
	{
		new-item $Monitoring_Host_Txt -type file -force	
	}
	

$Global:Monitoring_Host_Content = get-content $Monitoring_Host_Txt
If ($Monitoring_Host_Content -ne $null)
	{	
		$MonitoringHost = $Monitoring_Host_Content
		Try
			{
				$Web_Monitoring_URL = "http://" + $Monitoring_Host_Content + ":9801/MDTMonitorData/Computers/"
				$Data = Invoke-RestMethod $Web_Monitoring_URL
				$MonitoringHost_Status_Label.Content = "Connexion to Monitoring host OK"		
				$MonitoringHost_Status_Block.Background = "#00a300"							
			}
		Catch
			{
				$Menu_Open.Enabled = $false		
				$Menu_CSV.Enabled = $false	
				$MonitoringHost_Status_Label.Content = "Connexion to Monitoring host KO"		
				$MonitoringHost_Status_Block.Background = "Red"							
			}	
			
	}	
Else
	{
		$MonitoringHost_Status_Label.Content = "No monitoring host found"
		$MonitoringHost_Status_Block.Background = "Orange"					
	}


	
################################################################################################################################"
# MAIN FUNCTIONS
################################################################################################################################"

# ----------------------------------------------------------------------
# Part - Function to get the monitoring data from the server
# ----------------------------------------------------------------------	
	
function Get_My_MonitoringData
{ 
	param(
	$MyHost
	)
	$Web_Monitoring_URL = "http://" + $MyHost + ":9801/MDTMonitorData/Computers/"
	$Data = Invoke-RestMethod $Web_Monitoring_URL
  foreach($property in ($Data.content.properties)) 
  { 
		$Percent = $property.PercentComplete.'#text' 		
		$Current_Steps = $property.CurrentStep.'#text'			
		$Total_Steps = $property.TotalSteps.'#text'
		
		If ($Current_Steps -eq $Total_Steps)
			{
				If ($Percent -eq $null)
					{			
						$Step_Status = "Not started"
					}
				Else
					{
						$Step_Status = "$Current_Steps / $Total_Steps"
					}					
			}
		Else
			{
				$Step_Status = "$Current_Steps / $Total_Steps"			
			}
	
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
	  "DeployID" = $($property.UniqueID.'#text'); 			
	  "Computer Name" = $($property.Name); 
	  PercentNumber = $($property.PercentComplete.'#text');
	  "Percent Complete" = $Percent_Value; 	  
	  "Step Name" = $StepName;	  	  
	  "Step status" = $Step_Status;	  
	  Warnings = $($property.Warnings.'#text'); 
	  Errors = $($property.Errors.'#text'); 
	  DeployStatus = $($property.DeploymentStatus.'#text'	);		  
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
	  "Start time" = ([datetime]$($property.StartTime.'#text')).ToLocalTime().ToString('HH:mm:ss')  
	  "End time" = $EndTime;
	  "Ellapsed time" = $Ellapsed;	  	  
	} 
  } 
}	



# ----------------------------------------------------------------------
# Part - Function to populate the monitoring list 
# ----------------------------------------------------------------------	

Function Populate_Datagrid_Monitoring
	{		
		param(
		$MyHost
		)	
		
		$MyData = Get_My_MonitoringData -MyHost $MyHost | Select Date, "Computer Name", DeployID, "Deployment Status", "Percent Complete", PercentNumber, "Step Name", Warnings, Errors, "Start time", "End Time", "Ellapsed time", DeployStatus #| where {$_.ComputerName -eq $MyComputer}	
				
		$NB_Success = ($MyData | Where {$_."Deployment Status" -eq "Success"}).count
		$NB_Failed = ($MyData | Where {$_."Deployment Status" -eq "Failed"}).count
		$NB_Runnning = ($MyData | Where {$_."Deployment Status" -eq "Running"}).count
		$NB_Unresponsive = ($MyData | Where {$_."Deployment Status" -eq "Unresponsive"}).count			

		$Window_Running_Comp.content = "Running: $NB_Runnning"
		$Success_Comp.content = "Success: $NB_Success"		
		$Failed_Comp.content = "Failed: $NB_Failed"
		$Unresponsive_Comp.content = "Unresponsive: $NB_Unresponsive"

		If ($MyData -eq $null)
			{
			}
		Else
			{
				ForEach ($data in $MyData)				
					{
						$Monitor_values = New-Object PSObject
						$Monitor_values = $Monitor_values | Add-Member  Date $data.Date -passthru		
						$Monitor_values = $Monitor_values | Add-Member  DeploymentStatus $data."Deployment Status" -passthru							
						$Monitor_values = $Monitor_values | Add-Member  Name $data."Computer Name" -passthru
						$Monitor_values = $Monitor_values | Add-Member  Ellapsed_time $data."Ellapsed time" -passthru														
						$Monitor_values = $Monitor_values | Add-Member  PercentComplete $data."Percent Complete" -passthru														
						$Monitor_values = $Monitor_values | Add-Member  StepName $data."Step Name" -passthru	
						$DataGrid_Monitoring.Items.Add($Monitor_values) > $null		

					}
			}					
	}
	
	
# ----------------------------------------------------------------------
# Part - Function to populate the monitoring list of running computers
# ----------------------------------------------------------------------
	
# Populate the part be notified	when running deployment are finished
Function Populate_Running_Datagrid
	{		
		param(
		$MyHost
		)	
		
		$MyRunningData = Get_My_MonitoringData -MyHost $MyHost | Select Date, "Computer Name", DeployID, "Deployment Status", "Percent Complete", PercentNumber, "Step Name", Warnings, Errors, "Start time", "End Time", "Ellapsed time", DeployStatus |  where {$_."Deployment Status" -eq "Success"} 	

		If ($MyRunningData -eq $null)
			{
			}
		Else
			{
				ForEach ($data in $MyRunningData)				
					{
						$Running_values = New-Object PSObject
						$Running_values = $Running_values | Add-Member  Date $data.Date -passthru		
						$Running_values = $Running_values | Add-Member  DeploymentStatus $data."Deployment Status" -passthru							
						$Running_values = $Running_values | Add-Member  Name $data."Computer Name" -passthru
						$Running_values = $Running_values | Add-Member  Ellapsed_time $data."Ellapsed time" -passthru														
						$Running_values = $Running_values | Add-Member  PercentComplete $data."Percent Complete" -passthru														
						$Running_values = $Running_values | Add-Member  StepName $data."Step Name" -passthru	
						$DataGrid_Running.Items.Add($Running_values) > $null	
					}
			}					
	}	

	
	
	
	
################################################################################################################################"
# MAIN CONTROLS
################################################################################################################################"

# ----------------------------------------------------
# Part - Main Monitoring list
# ----------------------------------------------------		
	
	
# Create notifyicon, and right-click -> Exit menu
$MDTMonitoring_Icon = New-Object System.Windows.Forms.NotifyIcon
$MDTMonitoring_Icon.Text = "MDT Monitoring"
$MDTMonitoring_Icon.Icon = $icon
$MDTMonitoring_Icon.Visible = $true

$Menu_Open = New-Object System.Windows.Forms.MenuItem
$Menu_Open.Text = "Open monitoring list"

$Menu_Config = New-Object System.Windows.Forms.MenuItem
$Menu_Config.Text = "Monitoring Host settings"

$Menu_Notif = New-Object System.Windows.Forms.MenuItem
$Menu_Notif.Text = "Be notified when deployment are finished"

$Menu_CSV = New-Object System.Windows.Forms.MenuItem
$Menu_CSV.Text = "Export to a CSV report"

$Menu_About = New-Object System.Windows.Forms.MenuItem
$Menu_About.Text = "About"

$Tool_Gallery_Version = "1.0"
# $Tool_Gallery_Version = (find-script DeploymentMonitoring).version
If ($Current_Version -ne $Tool_Gallery_Version)
	{
		$Menu_Update = New-Object System.Windows.Forms.MenuItem
		$Menu_Update.Text = "New Update available !!!"		
	}

$Menu_Restart_Tool = New-Object System.Windows.Forms.MenuItem
$Menu_Restart_Tool.Text = "Restart the tool"

$Menu_Exit = New-Object System.Windows.Forms.MenuItem
$Menu_Exit.Text = "Exit"

$contextmenu = New-Object System.Windows.Forms.ContextMenu
$MDTMonitoring_Icon.ContextMenu = $contextmenu
$MDTMonitoring_Icon.contextMenu.MenuItems.AddRange($Menu_Open)
$MDTMonitoring_Icon.contextMenu.MenuItems.AddRange($Menu_Config)
$MDTMonitoring_Icon.contextMenu.MenuItems.AddRange($Menu_Notif)
$MDTMonitoring_Icon.contextMenu.MenuItems.AddRange($Menu_CSV)
$MDTMonitoring_Icon.contextMenu.MenuItems.AddRange($Menu_About)

If ($Current_Version -ne $Tool_Gallery_Version)
	{
		$MDTMonitoring_Icon.contextMenu.MenuItems.AddRange($Menu_Update)	

	}
	
$MDTMonitoring_Icon.contextMenu.MenuItems.AddRange($Menu_Restart_Tool)
$MDTMonitoring_Icon.contextMenu.MenuItems.AddRange($Menu_Exit)


# Timer to refresh the datagrid each 5 seconds
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 10000 
$timer.add_tick({UpdateUi})
 
Function UpdateUi()
{
	($DataGrid_Monitoring.items).Clear()	
	($DataGrid_Running.items).Clear()		
	Get_My_MonitoringData		
	$Global:Monitoring_Host_Content = get-content $Monitoring_Host_Txt
	Populate_Datagrid_Monitoring -MyHost $Monitoring_Host_Content 				
	# Populate_DataGrid_Monitoring
	Populate_Running_Datagrid -MyHost $Monitoring_Host_Content 	
}


################################################################################################################################"
# ACTIONS ON BUTTONS FROM THE CONTEXTMENU
################################################################################################################################"

# ---------------------------------------------------------------------
# Action when after a click on the systray icon
# ---------------------------------------------------------------------
$MDTMonitoring_Icon.Add_Click({
	$Monitoring_Host_Status = $false		

	If ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
		$window.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$window.Width)
		$window.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$window.Height)
		$window.Show()
		$window.Activate()	
	}			

	$Global:Monitoring_Host_Content = get-content $Monitoring_Host_Txt
	If ($Monitoring_Host_Content -eq $null)
		{
			$Monitoring_Host_Status = $false
			$DataGrid_Status_On.Visibility = "Collapsed"
			$DataGrid_Status_Off.Visibility = "Visible"	

			$Main_Monitoring_Status.Background = "Red"
			$Main_Monitoring_msg.Content = "No monitoring host found"	

			$Window_Running_Comp.content = "Running: 0"
			$Success_Comp.content = "Success: 0"		
			$Failed_Comp.content = "Failed: 0"
			$Unresponsive_Comp.content = "Unresponsive: 0"				
		}
	Else
		{
			Try
				{
					$Global:MyData = Get_My_MonitoringData -MyHost $Monitoring_Host_Content 			
					$Monitoring_Host_Status = $true	
				}
			Catch
				{
					$Monitoring_Host_Status = $false
					$DataGrid_Status_On.Visibility = "Collapsed"
					$DataGrid_Status_Off.Visibility = "Visible"
					
					$Main_Monitoring_Status.Background = "Red"
					$Main_Monitoring_msg.Content = "Can not reach the Monitoring host"						

					$Window_Running_Comp.content = "Running: 0"
					$Success_Comp.content = "Success: 0"		
					$Failed_Comp.content = "Failed: 0"
					$Unresponsive_Comp.content = "Unresponsive: 0"								
				}

			If 	($Monitoring_Host_Status -eq $true)
				{			
					If ($MyData -ne $null)
						{
							$Global:Timer_Status = $timer.Enabled
							If ($Timer_Status -eq $true)
								{
									$timer.Stop()	
									$timer.start()												
								}
							Else
								{
									$timer.start()			
								}															
								
							$Main_Monitoring_Status.Background = "#00a300"
							$Main_Monitoring_msg.Content = "Monitoring status is working"									
								
							$DataGrid_Monitoring.Items.clear()					
							$DataGrid_Status_On.Visibility = "Visible"
							$DataGrid_Status_Off.Visibility = "Collapsed"			
							Populate_Datagrid_Monitoring -MyHost $Monitoring_Host_Content 			
						}
					Else
						{
							$DataGrid_Status_On.Visibility = "Collapsed"
							$DataGrid_Status_Off.Visibility = "Visible"
							
							$Main_Monitoring_Status.Background = "Orange"
							$Main_Monitoring_msg.Content = "No computer in your Monitoring data"								

							$Window_Running_Comp.content = "Running: 0"
							$Success_Comp.content = "Success: 0"		
							$Failed_Comp.content = "Failed: 0"
							$Unresponsive_Comp.content = "Unresponsive: 0"							
						}							

					
				}
		}
})




# ---------------------------------------------------------------------
# Action after clicking on Open monitoring list
# ---------------------------------------------------------------------

$Menu_Open.Add_Click({
	$window.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$window.Width)
	$window.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$window.Height)
	$window.Show()
	$window.Activate()				

	$Monitoring_Host_Status = $false		

	If ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
			$window.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$window.Width)
			$window.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$window.Height)
			$window.Show()
			$window.Activate()	
	}			

	$Global:Monitoring_Host_Content = get-content $Monitoring_Host_Txt
	If ($Monitoring_Host_Content -eq $null)
		{
			$Monitoring_Host_Status = $false
			$DataGrid_Status_On.Visibility = "Collapsed"
			$DataGrid_Status_Off.Visibility = "Visible"	

			$Main_Monitoring_Status.Background = "Red"
			$Main_Monitoring_msg.Content = "No monitoring host found"	

			$Window_Running_Comp.content = "Running: 0"
			$Success_Comp.content = "Success: 0"		
			$Failed_Comp.content = "Failed: 0"
			$Unresponsive_Comp.content = "Unresponsive: 0"				
		}
	Else
		{
			Try
				{
					$Global:MyData = Get_My_MonitoringData -MyHost $Monitoring_Host_Content 			
					$Monitoring_Host_Status = $true	
				}
			Catch
				{
					$Monitoring_Host_Status = $false
					$DataGrid_Status_On.Visibility = "Collapsed"
					$DataGrid_Status_Off.Visibility = "Visible"
					
					$Main_Monitoring_Status.Background = "Red"
					$Main_Monitoring_msg.Content = "Can not reach the Monitoring host"						

					$Window_Running_Comp.content = "Running: 0"
					$Success_Comp.content = "Success: 0"		
					$Failed_Comp.content = "Failed: 0"
					$Unresponsive_Comp.content = "Unresponsive: 0"								
				}

			If 	($Monitoring_Host_Status -eq $true)
				{			
					If ($MyData -ne $null)
						{
							$Timer_Status = $timer.Enabled
							If ($Timer_Status -eq $true)
								{
									$timer.Stop()	
									$timer.start()												
								}
							Else
								{
									$timer.start()			
								}	
								
							$Main_Monitoring_Status.Background = "#00a300"
							$Main_Monitoring_msg.Content = "Monitoring status is working"									
								
							$DataGrid_Monitoring.Items.clear()					
							$DataGrid_Status_On.Visibility = "Visible"
							$DataGrid_Status_Off.Visibility = "Collapsed"			
							Populate_Datagrid_Monitoring -MyHost $Monitoring_Host_Content 			
						}
					Else
						{
							$DataGrid_Status_On.Visibility = "Collapsed"
							$DataGrid_Status_Off.Visibility = "Visible"
							
							$Main_Monitoring_Status.Background = "Orange"
							$Main_Monitoring_msg.Content = "No computer in your Monitoring data"								

							$Window_Running_Comp.content = "Running: 0"
							$Success_Comp.content = "Success: 0"		
							$Failed_Comp.content = "Failed: 0"
							$Unresponsive_Comp.content = "Unresponsive: 0"							
						}							

					
				}
		}
})




# ---------------------------------------------------------------------
# Action after clicking on Monitoring host settings
# ---------------------------------------------------------------------

$Menu_Config.Add_Click({	
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($window_Config)
	$window_Config.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$window_Config.Width)
	$window_Config.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$window_Config.Height)
	$window_Config.Show()
	$window_Config.Activate()	
	
	$Global:Monitoring_Host_Content = get-content $Monitoring_Host_Txt	
	If ($Monitoring_Host_Content -eq $null)
		{	
			$Monitoring_Host_Content = ""
			$Host_TxtBox.Text = ""			
			$MonitoringHost_Status_Label.Content = "No monitoring host found"
			$MonitoringHost_Status_Block.Background = "Orange"				
		}
	Else
		{
			$Host_TxtBox.Text = $Monitoring_Host_Content
			
			Try
				{
					$Global:MyData = Get_My_MonitoringData -MyHost $Monitoring_Host_Content 			
					$MonitoringHost_Status_Label.Content = "Connexion to Monitoring host OK"		
					$MonitoringHost_Status_Block.Background = "#00a300"							
				}
			Catch
				{
					$MonitoringHost_Status_Label.Content = "Connexion to Monitoring host KO"		
					$MonitoringHost_Status_Block.Background = "Red"						
				}
		}
})


# Menu Monitoring Host settings - Button set this host
$Set_Host.Add_Click({
	$MyHost = $Host_TxtBox.Text.ToString()
	If ($MyHost -ne "")
		{
			If (test-path $Monitoring_Host_Txt)
				{		
					Remove-item $Monitoring_Host_Txt -force	
				}
				
			new-item $Monitoring_Host_Txt -type file -force
			Add-content -path $Monitoring_Host_Txt -value $MyHost	
			
			Try
				{
					$Global:MyData = Get_My_MonitoringData -MyHost $MyHost 			
					$MonitoringHost_Status_Label.Content = "Connexion to Monitoring host OK"		
					$MonitoringHost_Status_Block.Background = "#00a300"							
				}
			Catch
				{
					$MonitoringHost_Status_Label.Content = "Connexion to Monitoring host KO"		
					$MonitoringHost_Status_Block.Background = "Red"						
				}				
		}
})


# ---------------------------------------------------------------------
# Action after clicking on Be notified when computers are installed
# ---------------------------------------------------------------------

$Menu_Notif.Add_Click({
	$Monitoring_Host_Status = $false	

	$Window_Part_Running_Comp.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$Window_Part_Running_Comp.Width)
	$Window_Part_Running_Comp.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$Window_Part_Running_Comp.Height)
	$Window_Part_Running_Comp.Show()
	$Window_Part_Running_Comp.Activate()

	$Global:Monitoring_Host_Content = get-content $Monitoring_Host_Txt
	If ($Monitoring_Host_Content -eq $null)
		{
			$Monitoring_Host_Status = $false
			$DataGrid_Running_On.Visibility = "Collapsed"
			$DataGrid_Running_Off.Visibility = "Visible"	
			
			$Main_Running_Monitoring_Status.Background = "Red"
			$Main_Running_Monitoring_msg.Content = "No monitoring host found"				

			$Window_Running_Comp.content = "Running: 0"
			$Success_Comp.content = "Success: 0"		
			$Failed_Comp.content = "Failed: 0"
			$Unresponsive_Comp.content = "Unresponsive: 0"				
		}
	Else
		{
			Try
				{
					$Global:MyData = Get_My_MonitoringData -MyHost $Monitoring_Host_Content 			
					$Monitoring_Host_Status = $true
				}
			Catch
				{
					$Monitoring_Host_Status = $false
					$DataGrid_Running_On.Visibility = "Collapsed"
					$DataGrid_Running_Off.Visibility = "Visible"
					
					$Main_Running_Monitoring_Status.Background = "Red"
					$Main_Running_Monitoring_msg.Content = "Can not reach the Monitoring host"					

					$Window_Running_Comp.content = "Running: 0"
					$Success_Comp.content = "Success: 0"		
					$Failed_Comp.content = "Failed: 0"
					$Unresponsive_Comp.content = "Unresponsive: 0"								
				}

			If 	($Monitoring_Host_Status -eq $true)
				{			
					If ($MyData -ne $null)
						{
							$Main_Running_Monitoring_Status.Background = "#00a300"
							$Main_Running_Monitoring_msg.Content = "Monitoring status is working"

							$DataGrid_Monitoring.Items.clear()					
							$DataGrid_Running_On.Visibility = "Visible"
							$DataGrid_Running_Off.Visibility = "Collapsed"			
							Populate_Running_Datagrid -MyHost $Monitoring_Host_Content 			
						}
					Else
						{
							$DataGrid_Running_On.Visibility = "Collapsed"
							$DataGrid_Running_Off.Visibility = "Visible"
							
							$Main_Running_Monitoring_Status.Background = "Orange"
							$Main_Running_Monitoring_msg.Content = "No running computers found"								

							$Window_Running_Comp.content = "Running: 0"
							$Success_Comp.content = "Success: 0"		
							$Failed_Comp.content = "Failed: 0"
							$Unresponsive_Comp.content = "Unresponsive: 0"							
						}							
				}
		}
})



# ---------------------------------------------------------------------
# Action after clicking on Export to a CSV report
# ---------------------------------------------------------------------

$Menu_CSV.Add_Click({
	$tmp_folder = $env:TEMP
	$Deployment_List = "$tmp_folder\Deployment_List.csv"
	If (test-path $Deployment_List)
		{
			remove-item $Deployment_List -force
		}
	$DataGrid_Monitoring.items | select DeploymentStatus, Date, Name, PercentComplete, Step_status, StepName, Warnings, Errors, Start_time, End_Time, Ellapsed_time, LastTime, DARTIP, DartPort, DartTicket, VMHost, VMName | export-csv $Deployment_List -NoTypeInformation	-UseCulture		
	invoke-item $Deployment_List
})



# ---------------------------------------------------------------------
# Action after clicking on About
# ---------------------------------------------------------------------

$Menu_About.Add_Click({
	$window_About.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$window_About.Width)
	$window_About.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$window_About.Height)
	$window_About.Show()
	$window_About.Activate()
	
	If ($Current_Version -lt $Tool_Gallery_Version)
		{
			$Main_Update_Status.Background = "Orange"
			$Update_Status_Label.Content = "A new version is available"
			$Update_Status_Label.Foreground = "White"
			$Update_Status_Label.FontWeight="Bold"
			$Update_Status_Label.Fontsize = "12"
			$Update_Status_Label.Margin = "0,7,0,0"		
		}
})










################################################################################################################################"
# ACTIONS ON THE DIFFERENT WINDOW FROM CONTEXT MENU
################################################################################################################################"

# ----------------------------------------------------
# Part - Open Monitoring list
# ----------------------------------------------------

# Close the window if it's double clicked
$window.Add_MouseDoubleClick({
	$window.Hide()
	
	$Global:Timer_Status = $timer.Enabled
	If ($Timer_Status -eq $true)
		{
			$timer.Stop()	
		}					
})

# Close the window if it loses focus
$window.Add_Deactivated({
	$window.Hide()
	
	$Global:Timer_Status = $timer.Enabled
	If ($Timer_Status -eq $true)
		{
			$timer.Stop()	
		}					
})



# ----------------------------------------------------
# Part - Monitoring host settings
# ----------------------------------------------------

# Close the window if it's double clicked
$window_Config.Add_MouseDoubleClick({
	$window_Config.Hide()
})

# Close the window if it loses focus
$window_Config.Add_Deactivated({
	$window_Config.Hide()
})


# ----------------------------------------------------
# Part - Be notified when computers are installed
# ----------------------------------------------------

# Close the window if it's double clicked
$Window_Part_Running_Comp.Add_MouseDoubleClick({
	$Window_Part_Running_Comp.Hide()
	
	$Global:Timer_Status = $timer.Enabled
	If ($Timer_Status -eq $true)
		{
			$timer.Stop()	
		}					
})

# Close the window if it loses focus
$Window_Part_Running_Comp.Add_Deactivated({
	$Window_Part_Running_Comp.Hide()
	
	$Global:Timer_Status = $timer.Enabled
	If ($Timer_Status -eq $true)
		{
			$timer.Stop()	
		}					
})

$NotifyMe.Add_Click({

	ForEach ($Selected_item in $DataGrid_Running.Selecteditems)
		{
			$Global:Selected_item_Name = $Selected_item.Name	
		}

	start-process powershell.exe ".\MonitoringData_Deploy_Analyze.ps1 '$Selected_item_Name' '$Monitoring_Host_Content'" 							
})






# ----------------------------------------------------
# Part - About Window
# ----------------------------------------------------

# Close the window if it's double clicked
$window_About.Add_MouseDoubleClick({
	$window_About.Hide()
})

# Close the window if it loses focus
$window_About.Add_Deactivated({
	$window_About.Hide()
})


# ----------------------------------------------------
# Part - Exit
# ----------------------------------------------------

# When Exit is clicked, close everything and kill the PowerShell process
$Menu_Restart_Tool.add_Click({
	$Restart = "Yes"
	start-process -WindowStyle hidden powershell.exe ".\MDT_Systray_Solution.ps1 '$Restart'" 	

	$MDTMonitoring_Icon.Visible = $false
	$window.Close()
	# $window_Config.Close()	
	Stop-Process $pid
	
	$Global:Timer_Status = $timer.Enabled
	If ($Timer_Status -eq $true)
		{
			$timer.Stop()	
		}		
 })
 

# ----------------------------------------------------
# Part - Exit
# ----------------------------------------------------

# When Exit is clicked, close everything and kill the PowerShell process
$Menu_Exit.add_Click({
	$MDTMonitoring_Icon.Visible = $false
	$window.Close()
	# $window_Config.Close()	
	Stop-Process $pid
	
	$Global:Timer_Status = $timer.Enabled
	If ($Timer_Status -eq $true)
		{
			$timer.Stop()	
		}	
 })
 
 
 

# Make PowerShell Disappear
$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

# Force garbage collection just to start slightly lower RAM usage.
[System.GC]::Collect()



# Create an application context for it to all run within.
# This helps with responsiveness, especially when clicking Exit.
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)