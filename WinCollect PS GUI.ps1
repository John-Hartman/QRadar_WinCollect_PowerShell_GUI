########################################################################################################################### 
#                  Powershell GUI Tool for WinCollect Troubleshooting  ###############
# Created by John Hartman   Email=john.hartman@ibm.com
# Version:1
# 
#
#
# Instructions for use:
#	
#	WinCollect Service Buttons
#	
#		Check - checks the status of the WinCollect service on that computer
#		Start - starts the WinCollect service on that computer
#		Stop - stops the WinCollect service on that computer
#		
#	Port Check Buttons
#		
#		1) Enter the IP of the QRadar destination
#		2) Enter all IPs of the systems that are being polled by or are forwarding to this server
#
#		Port Check - with a destination server entered into the "Enter Servers" box you can type any port number into the "Filters" box and test the connectivity (only works for TCP)
#		8413 - With the IP of the QRadar managed host entered in the "Enter Servers" box this will test the conectivity with that configuration console
#		135 - With the IP of the windows systems that are forwarding to or being polled by this WinCollect server, this will test their port connectivity
#		139 - With the IP of the windows systems that are forwarding to or being polled by this WinCollect server, this will test their port connectivity
#		445 - With the IP of the windows systems that are forwarding to or being polled by this WinCollect server, this will test their port connectivity
#	
#	WinCollect Log Buttons
#		
#		Open - Opens the WinCollect Log file
#		Search - Type a search filter e.g. IP address, or any string into the "Filters" box and this button will return those lines in the "Output" box
#		
#	Install Config Buttons
#		
#		Open - Opens the Install Config file
#		CHK Identifier - Returns the Identifier line from the Install Config file into the "Output" box
#		CHG Identifier - Type the Current Identifier value in the "Enter Servers" box and the new desired value in the "Filters" box to change the Identifier
#		CHK Server - Returns the Status Server line from the Install Config file into the "Output" box
#		CHG Server - Type the Current Status Server value in the "Enter Servers" box and the new desired value in the "Filters" box to change the Status Server and Configuration Server
#	
#	New PEM button
#	
#		Generate - First click the "Stop" button under the WinCollect Service section then press the "Generate" button which will append ".old" to the end of the PEM file. then click the "Start" button under the WinCollect Service section which will create a new PEM file.
#		
#
#  
#
#    Note: Please copy  "Lucida Sans Typewriter,9"  font in your server where
#    this tool is running in order to get the out put in clearly
#
#
#                              
########################################################################################################################### 
 
 
# region Form 
 
Add-Type -AssemblyName System.Windows.Forms 
 
$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "WinCollect Troubleshooting - Created by John Hartman" 
$Form.TopMost = $true 
$Form.Width = 575
$Form.Height = 700 
$Form.FormBorderStyle= "Sizable" 
$form.StartPosition ="centerScreen" 
$form.ShowInTaskbar = $true 
$form.BackColor = "#101720" 
$form.HelpButton = $true

    
$StatusBar = New-Object System.Windows.Forms.StatusBar
$StatusBar.Text = "Ready"
$StatusBar.Height = 22
$StatusBar.Width = 200
$StatusBar.Location = New-Object System.Drawing.Point( 0, 250 )
$Form.Controls.Add($StatusBar)

 
# endregion

 
# region Text Boxes
 
$InputBox = New-Object system.windows.Forms.TextBox 
$InputBox.Multiline = $true 
$InputBox.BackColor = "#18212D" 
$InputBox.Width = 280 
$InputBox.Height = 135 
$InputBox.ScrollBars ="Vertical" 
$InputBox.location = new-object system.drawing.point(250,30) 
$InputBox.Font = "Microsoft Sans Serif,10,style=Bold" 
$InputBox.ForeColor = "#FFFFFF"
$Form.controls.Add($inputbox) 

 
$filterbox= New-Object system.windows.Forms.TextBox 
$filterbox.Multiline = $true 
$filterBox.BackColor = "#18212D" 
$filterbox.Width = 280 
$filterbox.Height = 135
$filterbox.ScrollBars ="Vertical" 
$filterbox.location = new-object system.drawing.point(250,190) 
$filterbox.Font = "Microsoft Sans Serif,10" 
$filterbox.ForeColor = "#FFFFFF"
$Form.controls.Add($filterbox) 

 
$outputBox= New-Object System.Windows.Forms.RichTextBox 
$outputBox.Multiline = $true 
$outputBox.BackColor = "#18212D" 
$outputBox.Width = 500
$outputBox.Height = 265
$outputBox.ReadOnly =$true 
$outputBox.ScrollBars = "Both" 
$outputBox.WordWrap = $false 
$outputBox.location = new-object system.drawing.point(10,350) 
$outputBox.Font = "Lucida Sans Typewriter,9" 
$outputBox.ForeColor = "#FFFFFF"
$Form.controls.Add($outputBox) 
 
  


# endregion


# region Labels

$label3 = New-Object system.windows.Forms.Label 
$label3.Text = "Enter Servers" 
$label3.AutoSize = $true 
$label3.Width = 25 
$label3.Height = 10 
$label3.location = new-object system.drawing.point(250,10) 
$label3.Font = "Microsoft Sans Serif,10,style=Bold" 
$label3.ForeColor = "#FFFFFF"
$Form.controls.Add($label3) 
 
 
$Filters = New-Object system.windows.Forms.Label 
$Filters.Text = "Filters" 
$Filters.AutoSize = $true 
$Filters.Width = 25 
$Filters.Height = 10 
$Filters.location = new-object system.drawing.point(250,170) 
$Filters.Font = "Microsoft Sans Serif,10,style=Bold"
$Filters.ForeColor = "#FFFFFF"
$Form.controls.Add($Filters) 
 
 
$Portchk = New-Object system.windows.Forms.Label 
$Portchk.Text = "Port Check" 
$Portchk.AutoSize = $true 
$Portchk.Width = 25 
$Portchk.Height = 10 
$Portchk.location = new-object system.drawing.point(10,100) 
$Portchk.Font = "Microsoft Sans Serif,10,style=Bold" 
$Portchk.ForeColor = "#FFFFFF"
$Form.controls.Add($Portchk)
 
 
$wcsvc = New-Object system.windows.Forms.Label 
$wcsvc.Text = "WinCollect Svc" 
$wcsvc.AutoSize = $true 
$wcsvc.Width = 25 
$wcsvc.Height = 10 
$wcsvc.location = new-object system.drawing.point(10,10) 
$wcsvc.Font = "Microsoft Sans Serif,10,style=Bold" 
$wcsvc.ForeColor = "#FFFFFF"
$Form.controls.Add($wcsvc)


$Outputlb = New-Object system.windows.Forms.Label 
$Outputlb.Text = "Output" 
$Outputlb.AutoSize = $true 
$Outputlb.Width = 25 
$Outputlb.Height = 10 
$Outputlb.location = new-object system.drawing.point(10,330) 
$Outputlb.Font = "Microsoft Sans Serif,10,style=Bold"
$Outputlb.ForeColor = "#FFFFFF"
$Form.controls.Add($Outputlb) 


$WCloglb = New-Object system.windows.Forms.Label 
$WCloglb.Text = "WinCollect Log" 
$WCloglb.AutoSize = $true 
$WCloglb.Width = 25 
$WCloglb.Height = 10 
$WCloglb.location = new-object system.drawing.point(130,10) 
$WCloglb.Font = "Microsoft Sans Serif,10,style=Bold"
$WCloglb.ForeColor = "#FFFFFF"
$Form.controls.Add($WCloglb)


$InstallConflb = New-Object system.windows.Forms.Label 
$InstallConflb.Text = "Install Config" 
$InstallConflb.AutoSize = $true 
$InstallConflb.Width = 25 
$InstallConflb.Height = 10 
$InstallConflb.location = new-object system.drawing.point(130,100) 
$InstallConflb.Font = "Microsoft Sans Serif,10,style=Bold"
$InstallConflb.ForeColor = "#FFFFFF"
$Form.controls.Add($InstallConflb)


$PEMFilelb = New-Object system.windows.Forms.Label 
$PEMFilelb.Text = "New PEM" 
$PEMFilelb.AutoSize = $true 
$PEMFilelb.Width = 25 
$PEMFilelb.Height = 10 
$PEMFilelb.location = new-object system.drawing.point(10,230) 
$PEMFilelb.Font = "Microsoft Sans Serif,10,style=Bold"
$PEMFilelb.ForeColor = "#FFFFFF"
$Form.controls.Add($PEMFilelb)


# endregion


##########    Buttons    ##########

# region WinCollect Service Buttons
  

$CheckSerbutton = New-Object system.windows.Forms.Button 
$CheckSerbutton.BackColor = "#66ABAD" 
$CheckSerbutton.Text = "Check" 
$CheckSerbutton.Width = 100 
$CheckSerbutton.Height = 22
$CheckSerbutton.location = new-object system.drawing.point(10,30) 
$CheckSerbutton.Font = "Microsoft Sans Serif,8" 
$CheckSerbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$CheckSerbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$CheckSerbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$CheckSerbutton.Add_Click({Check-ser}) 
$Form.controls.Add($CheckSerbutton) 
  
$StartSerbutton = New-Object system.windows.Forms.Button 
$StartSerbutton.BackColor = "#66ABAD" 
$StartSerbutton.Text = "Start" 
$StartSerbutton.Width = 100 
$StartSerbutton.Height = 22
$StartSerbutton.location = new-object system.drawing.point(10,50) 
$StartSerbutton.Font = "Microsoft Sans Serif,8" 
$StartSerbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$StartSerbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$StartSerbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$StartSerbutton.Add_Click({Start-ser}) 
$Form.controls.Add($StartSerbutton) 
  
  
$StopSerbutton = New-Object system.windows.Forms.Button 
$StopSerbutton.BackColor = "#66ABAD" 
$StopSerbutton.Text = "Stop" 
$StopSerbutton.Width = 100 
$StopSerbutton.Height = 22
$StopSerbutton.location = new-object system.drawing.point(10,70) 
$StopSerbutton.Font = "Microsoft Sans Serif,8" 
$StopSerbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$StopSerbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$StopSerbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$StopSerbutton.Add_Click({Stop-ser}) 
$Form.controls.Add($StopSerbutton) 
  
  
# endregion
  

# region Port Check Buttons  
  
$portbutton = New-Object system.windows.Forms.Button 
$portbutton.BackColor = "#66ABAD" 
$portbutton.Text = "Port check" 
$portbutton.Width = 100
$portbutton.Height = 22
$portbutton.location = new-object system.drawing.point(10,120) 
$portbutton.Font = "Microsoft Sans Serif,8" 
$portbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$portbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$portbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$portbutton.Add_Click({Get-portstatus}) 
$Form.controls.Add($portbutton) 
 
 
$portbutton = New-Object system.windows.Forms.Button 
$portbutton.BackColor = "#66ABAD" 
$portbutton.Text = "8413" 
$portbutton.Width = 100
$portbutton.Height = 22
$portbutton.location = new-object system.drawing.point(10,140) 
$portbutton.Font = "Microsoft Sans Serif,8" 
$portbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$portbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$portbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$portbutton.Add_Click({Get-8413portstatus}) 
$Form.controls.Add($portbutton) 

 
$portbutton = New-Object system.windows.Forms.Button 
$portbutton.BackColor = "#66ABAD" 
$portbutton.Text = "135" 
$portbutton.Width = 100
$portbutton.Height = 22
$portbutton.location = new-object system.drawing.point(10,160) 
$portbutton.Font = "Microsoft Sans Serif,8" 
$portbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$portbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$portbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$portbutton.Add_Click({Get-135portstatus}) 
$Form.controls.Add($portbutton)


$portbutton = New-Object system.windows.Forms.Button 
$portbutton.BackColor = "#66ABAD" 
$portbutton.Text = "139" 
$portbutton.Width = 100
$portbutton.Height = 22
$portbutton.location = new-object system.drawing.point(10,180) 
$portbutton.Font = "Microsoft Sans Serif,8" 
$portbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$portbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$portbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$portbutton.Add_Click({Get-139portstatus}) 
$Form.controls.Add($portbutton)


$portbutton = New-Object system.windows.Forms.Button 
$portbutton.BackColor = "#66ABAD" 
$portbutton.Text = "445" 
$portbutton.Width = 100
$portbutton.Height = 22
$portbutton.location = new-object system.drawing.point(10,200) 
$portbutton.Font = "Microsoft Sans Serif,8" 
$portbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$portbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$portbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$portbutton.Add_Click({Get-445portstatus}) 
$Form.controls.Add($portbutton)
 
 
# endregion
 
 
# region WinCollect log Buttons


$WClogOpenbutton = New-Object system.windows.Forms.Button 
$WClogOpenbutton.BackColor = "#66ABAD" 
$WClogOpenbutton.Text = "Open" 
$WClogOpenbutton.Width = 100 
$WClogOpenbutton.Height = 22 
$WClogOpenbutton.location = new-object system.drawing.point(130,30) 
$WClogOpenbutton.Font = "Microsoft Sans Serif,8" 
$WClogOpenbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$WClogOpenbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$WClogOpenbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$WClogOpenbutton.Add_Click({WCLogopen}) 
$Form.controls.Add($WClogOpenbutton)

$WClogSearchbutton = New-Object system.windows.Forms.Button 
$WClogSearchbutton.BackColor = "#66ABAD" 
$WClogSearchbutton.Text = "Search" 
$WClogSearchbutton.Width = 100 
$WClogSearchbutton.Height = 22 
$WClogSearchbutton.location = new-object system.drawing.point(130,50) 
$WClogSearchbutton.Font = "Microsoft Sans Serif,8" 
$WClogSearchbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$WClogSearchbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$WClogSearchbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$WClogSearchbutton.Add_Click({WCsearch}) 
$Form.controls.Add($WClogSearchbutton)


# endregion


# region Install Config Buttons


$InstallOpenbutton = New-Object system.windows.Forms.Button 
$InstallOpenbutton.BackColor = "#66ABAD" 
$InstallOpenbutton.Text = "Open" 
$InstallOpenbutton.Width = 100 
$InstallOpenbutton.Height = 22 
$InstallOpenbutton.location = new-object system.drawing.point(130,120) 
$InstallOpenbutton.Font = "Microsoft Sans Serif,8" 
$InstallOpenbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$InstallOpenbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$InstallOpenbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$InstallOpenbutton.Add_Click({InstallConfOpen}) 
$Form.controls.Add($InstallOpenbutton)


$CHKIDbutton = New-Object system.windows.Forms.Button 
$CHKIDbutton.BackColor = "#66ABAD" 
$CHKIDbutton.Text = "CHK Identifier" 
$CHKIDbutton.Width = 100 
$CHKIDbutton.Height = 22 
$CHKIDbutton.location = new-object system.drawing.point(130,140) 
$CHKIDbutton.Font = "Microsoft Sans Serif,8" 
$CHKIDbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$CHKIDbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$CHKIDbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$CHKIDbutton.Add_Click({CHKID}) 
$Form.controls.Add($CHKIDbutton)


$CHGIDbutton = New-Object system.windows.Forms.Button 
$CHGIDbutton.BackColor = "#66ABAD" 
$CHGIDbutton.Text = "CHG Identifier" 
$CHGIDbutton.Width = 100 
$CHGIDbutton.Height = 22 
$CHGIDbutton.location = new-object system.drawing.point(130,160) 
$CHGIDbutton.Font = "Microsoft Sans Serif,8" 
$CHGIDbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$CHGIDbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$CHGIDbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$CHGIDbutton.Add_Click({CHGID}) 
$Form.controls.Add($CHGIDbutton)


$CHKSERbutton = New-Object system.windows.Forms.Button 
$CHKSERbutton.BackColor = "#66ABAD" 
$CHKSERbutton.Text = "CHK Server" 
$CHKSERbutton.Width = 100 
$CHKSERbutton.Height = 22 
$CHKSERbutton.location = new-object system.drawing.point(130,180) 
$CHKSERbutton.Font = "Microsoft Sans Serif,8" 
$CHKSERbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$CHKSERbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$CHKSERbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$CHKSERbutton.Add_Click({CHKSER}) 
$Form.controls.Add($CHKSERbutton)


$CHGSERbutton = New-Object system.windows.Forms.Button 
$CHGSERbutton.BackColor = "#66ABAD" 
$CHGSERbutton.Text = "CHG Server" 
$CHGSERbutton.Width = 100 
$CHGSERbutton.Height = 22 
$CHGSERbutton.location = new-object system.drawing.point(130,200) 
$CHGSERbutton.Font = "Microsoft Sans Serif,8" 
$CHGSERbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$CHGSERbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$CHGSERbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$CHGSERbutton.Add_Click({CHGSER}) 
$Form.controls.Add($CHGSERbutton)


# endregion


# region PEM File Buttons


$Generatebutton = New-Object system.windows.Forms.Button 
$Generatebutton.BackColor = "#66ABAD" 
$Generatebutton.Text = "Generate" 
$Generatebutton.Width = 100 
$Generatebutton.Height = 22 
$Generatebutton.location = new-object system.drawing.point(10,250) 
$Generatebutton.Font = "Microsoft Sans Serif,8" 
$Generatebutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$Generatebutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Generatebutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$Generatebutton.Add_Click({NEWPEM}) 
$Form.controls.Add($Generatebutton)


# endregion


# region Documentation Button


$Documentationbutton = New-Object system.windows.Forms.Button 
$Documentationbutton.BackColor = "#66ABAD" 
$Documentationbutton.Text = "Documentation" 
$Documentationbutton.Width = 100 
$Documentationbutton.Height = 22 
$Documentationbutton.location = new-object system.drawing.point(130,250) 
$Documentationbutton.Font = "Microsoft Sans Serif,8" 
$Documentationbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 255, 250)
$Documentationbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Documentationbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$Documentationbutton.Add_Click({SHOWDOC}) 
$Form.controls.Add($Documentationbutton)


# endregion

 
##########    Functions    ##########
 
# region Progress Display Function

Function Progressbar 
{ 
Add-Type -AssemblyName system.windows.forms 
$Script:formt = New-Object System.Windows.Forms.Form 
$Script:formt.Text = 'Please Wait' 
$Script:formt.TopMost = $true 
$Script:formt.StartPosition ="CenterScreen" 
$Script:formt.Width = 500 
$Script:formt.Height = 20 
$Script:formt.MaximizeBox = $false 
$Script:formt.MinimizeBox = $false 
$Script:formt.Visible = $false 
 
 
} 

# endregion
   

# region Check Service Function
 
function Check-ser { 
 
progressbar 
 $outputBox.Clear() 
 $statusBar.Text=("Processing the request")
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
 $Script:formt.text="Working on $computer" 
$infser +=  Get-service WinCollect  
$sl = $infser| ft -AutoSize |Out-String 
 $outputBox.Appendtext("{0}`n" -f $sl +"`n $ct"  )  
 $statusBar.Text=("Ready")
 $Script:formt.close()  
 } 
    
# endregion
 

# region Start Service Function


function Start-ser { 
 
progressbar 
 $outputBox.Clear() 
 $statusBar.Text=("Processing the request")
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
 $Script:formt.text="Working on $computer" 
$infser +=  Start-service WinCollect  
$sl = $infser| ft -AutoSize |Out-String 
 $outputBox.Appendtext("{0}`n" -f $sl +"`n $ct"  )  
 $statusBar.Text=("Ready")
 $Script:formt.close()  
 } 
 

# endregion


# region Stop Service Function

function Stop-ser { 
 
progressbar 
 $outputBox.Clear() 
 $statusBar.Text=("Processing the request")
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
 $Script:formt.text="Working on $computer" 
$infser +=  Stop-service WinCollect  
$sl = $infser| ft -AutoSize |Out-String 
 $outputBox.Appendtext("{0}`n" -f $sl +"`n $ct"  )  
 $statusBar.Text=("Ready")
 $Script:formt.close()  
 } 

# endregion
 
 
# region Port Check Function
 
function checkport { 
 param( 
 $computername =$env:computername 
 ) 
  $sname =$filterbox.text 
 $os = Test-NetConnection -ComputerName $computername -port $sname -ea silentlycontinue 
 if($os){ 
 
 $TcpTestSucceeded =$os.TcpTestSucceeded 
 
 $servername=$os.ComputerName 
  
 
  
 
 $results =new-object psobject 
 
 $results |Add-Member noteproperty TcpTestSucceeded  $TcpTestSucceeded 
 $results |Add-Member noteproperty ComputerName  $servername 
  
 
 
 #Display the results 
 
 $results | Select-Object computername,TcpTestSucceeded 
 
 } 
 
 
 else 
 
 { 
 
 $results =New-Object psobject 
 
 $results =new-object psobject 
 $results |Add-Member noteproperty TcpTestSucceeded "Na" 
 $results |Add-Member noteproperty ComputerName $servername 
 
 
  
 #display the results 
 
 $results | Select-Object computername,TcpTestSucceeded 
 
 
 
 
 } 
 
 
 
 } 
 
 $infoport =@() 
 
 
 foreach($allserver in $allservers){ 
 
$infoport += checkport $allserver  
 } 
 
 $infoport 
 
 
function Get-portstatus { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request")
 $computers=$InputBox.lines.Split("`n") 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
  $infoport =@() 
 foreach ($computer in $computers) 
 { 
  $Script:formt.text="Working on $computer" 
 $infoport +=  checkport $computer  
 $pres=  $infoport| ft -AutoSize  | Out-String 
  } 
 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 
 
 
    
 
# endregion


# region 8413 Port Check Function


function checkport8413 { 
 param( 
 $computername =$env:computername 
 ) 
  $sname =8413 
 $os = Test-NetConnection -ComputerName $computername -port $sname -ea silentlycontinue 
 if($os){ 
 
 $TcpTestSucceeded =$os.TcpTestSucceeded 
 
 $servername=$os.ComputerName 
  
 
  
 
 $results =new-object psobject 
 
 $results |Add-Member noteproperty TcpTestSucceeded  $TcpTestSucceeded 
 $results |Add-Member noteproperty ComputerName  $servername 
  
 
 
 #Display the results 
 
 $results | Select-Object computername,TcpTestSucceeded 
 
 } 
 
 
 else 
 
 { 
 
 $results =New-Object psobject 
 
 $results =new-object psobject 
 $results |Add-Member noteproperty TcpTestSucceeded "Na" 
 $results |Add-Member noteproperty ComputerName $servername 
 
 
  
 #display the results 
 
 $results | Select-Object computername,TcpTestSucceeded 
 
 
 
 
 } 
 
 
 
 } 
 
 $infoport =@() 
 
 
 foreach($allserver in $allservers){ 
 
$infoport += checkport $allserver  
 } 
 
 $infoport 
 
 
function Get-8413portstatus { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request")
 $computers=$InputBox.lines.Split("`n") 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
  $infoport =@() 
 foreach ($computer in $computers) 
 { 
  $Script:formt.text="Working on $computer" 
 $infoport +=  checkport8413 $computer  
 $pres=  $infoport| ft -AutoSize  | Out-String 
  } 
 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 
 

# endregion


# region 135 Port Check Function


function checkport135 { 
 param( 
 $computername =$env:computername 
 ) 
  $sname =135 
 $os = Test-NetConnection -ComputerName $computername -port $sname -ea silentlycontinue 
 if($os){ 
 
 $TcpTestSucceeded =$os.TcpTestSucceeded 
 
 $servername=$os.ComputerName 
  
 
  
 
 $results =new-object psobject 
 
 $results |Add-Member noteproperty TcpTestSucceeded  $TcpTestSucceeded 
 $results |Add-Member noteproperty ComputerName  $servername 
  
 
 
 #Display the results 
 
 $results | Select-Object computername,TcpTestSucceeded 
 
 } 
 
 
 else 
 
 { 
 
 $results =New-Object psobject 
 
 $results =new-object psobject 
 $results |Add-Member noteproperty TcpTestSucceeded "Na" 
 $results |Add-Member noteproperty ComputerName $servername 
 
 
  
 #display the results 
 
 $results | Select-Object computername,TcpTestSucceeded 
 
 
 
 
 } 
 
 
 
 } 
 
 $infoport =@() 
 
 
 foreach($allserver in $allservers){ 
 
$infoport += checkport $allserver  
 } 
 
 $infoport 
 
 
function Get-135portstatus { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request")
 $computers=$InputBox.lines.Split("`n") 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
  $infoport =@() 
 foreach ($computer in $computers) 
 { 
  $Script:formt.text="Working on $computer" 
 $infoport +=  checkport135 $computer  
 $pres=  $infoport| ft -AutoSize  | Out-String 
  } 
 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 


# endregion


# region 139 Port Check Function


function checkport139 { 
 param( 
 $computername =$env:computername 
 ) 
  $sname =139 
 $os = Test-NetConnection -ComputerName $computername -port $sname -ea silentlycontinue 
 if($os){ 
 
 $TcpTestSucceeded =$os.TcpTestSucceeded 
 
 $servername=$os.ComputerName 
  
 
  
 
 $results =new-object psobject 
 
 $results |Add-Member noteproperty TcpTestSucceeded  $TcpTestSucceeded 
 $results |Add-Member noteproperty ComputerName  $servername 
  
 
 
 #Display the results 
 
 $results | Select-Object computername,TcpTestSucceeded 
 
 } 
 
 
 else 
 
 { 
 
 $results =New-Object psobject 
 
 $results =new-object psobject 
 $results |Add-Member noteproperty TcpTestSucceeded "Na" 
 $results |Add-Member noteproperty ComputerName $servername 
 
 
  
 #display the results 
 
 $results | Select-Object computername,TcpTestSucceeded 
 
 
 
 
 } 
 
 
 
 } 
 
 $infoport =@() 
 
 
 foreach($allserver in $allservers){ 
 
$infoport += checkport $allserver  
 } 
 
 $infoport 
 
 
function Get-139portstatus { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request")
 $computers=$InputBox.lines.Split("`n") 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
  $infoport =@() 
 foreach ($computer in $computers) 
 { 
  $Script:formt.text="Working on $computer" 
 $infoport +=  checkport139 $computer  
 $pres=  $infoport| ft -AutoSize  | Out-String 
  } 
 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 


# endregion


# region 445 Port Check Function


function checkport445 { 
 param( 
 $computername =$env:computername 
 ) 
  $sname =445 
 $os = Test-NetConnection -ComputerName $computername -port $sname -ea silentlycontinue 
 if($os){ 
 
 $TcpTestSucceeded =$os.TcpTestSucceeded 
 
 $servername=$os.ComputerName 
  
 
  
 
 $results =new-object psobject 
 
 $results |Add-Member noteproperty TcpTestSucceeded  $TcpTestSucceeded 
 $results |Add-Member noteproperty ComputerName  $servername 
  
 
 
 #Display the results 
 
 $results | Select-Object computername,TcpTestSucceeded 
 
 } 
 
 
 else 
 
 { 
 
 $results =New-Object psobject 
 
 $results =new-object psobject 
 $results |Add-Member noteproperty TcpTestSucceeded "Na" 
 $results |Add-Member noteproperty ComputerName $servername 
 
 
  
 #display the results 
 
 $results | Select-Object computername,TcpTestSucceeded 
 
 
 
 
 } 
 
 
 
 } 
 
 $infoport =@() 
 
 
 foreach($allserver in $allservers){ 
 
$infoport += checkport $allserver  
 } 
 
 $infoport 
 
 
function Get-445portstatus { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request")
 $computers=$InputBox.lines.Split("`n") 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
  $infoport =@() 
 foreach ($computer in $computers) 
 { 
  $Script:formt.text="Working on $computer" 
 $infoport +=  checkport445 $computer  
 $pres=  $infoport| ft -AutoSize  | Out-String 
  } 
 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 
 
 


# endregion


# region New PEM Generation Function


function new-pem {
	
Rename-Item -Path "C:\Program Files\IBM\WinCollect\config\ConfigurationServer.PEM" -NewName "ConfigurationServer.PEM.old"
		
}


function NEWPEM { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request")
 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 

  $Script:formt.text="Working on $nval" 
 $newpem +=  new-pem  
 $pres=  $newpem| ft -AutoSize  | Out-String 

 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 
 
 
# endregion


# region WinCollect Log Open Function


function WCLogOpen { 

start 'C:\Program Files\IBM\WinCollect\logs\WinCollect.log'

}


# endregion


# region WinCollect Log Search Function

function filter-search { 
 
$sname =$filterbox.text 
Select-String -path 'C:\Program Files\IBM\WinCollect\logs\WinCollect.log' -pattern "$sname"
 
}

function WCsearch { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request")
 $snames=$filterbox.lines.Split("`n") 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
  $infoport =@() 
 foreach ($sname in $snames) 
 { 
  $Script:formt.text="Working on $sname" 
 $infoport +=  filter-search $sname  
 $pres=  $infoport| ft -AutoSize  | Out-String 
  } 
 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 
 
 
# endregion


# region Install Config Open Function


function InstallConfOpen { 

start 'C:\Program Files\IBM\WinCollect\config\install_config.txt'

}



# endregion


# region Install Config CHK Identifier Function


function identifier-check { 
 
Select-String -path 'C:\Program Files\IBM\WinCollect\config\install_config.txt' -pattern "ApplicationIdentifier"
 
}

function CHKID { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request") 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
 
  $Script:formt.text="Working" 
 $infoport +=  identifier-check 
 $pres=  $infoport| ft -AutoSize  | Out-String 
 
 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 


# endregion


# region Install Config CHG Identifier Function


function identifier-change { 

$oval =$inputbox.text
$nval =$filterbox.text
((Get-Content -path "C:\Program Files\IBM\WinCollect\config\install_config.txt" -Raw) -replace $oval,$nval) | Set-Content -Path "C:\Program Files\IBM\WinCollect\config\install_config.txt"
 
}

function CHGID { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request")
 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 

  $Script:formt.text="Working on $nval" 
 $idchange +=  identifier-change  
 $pres=  $idchange| ft -AutoSize  | Out-String 

 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 



# endregion


# region Install Config CHK Server Function


function server-check { 
 
Select-String -path 'C:\Program Files\IBM\WinCollect\config\install_config.txt' -pattern "StatusServer"
 
}

function CHKSER { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request") 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 
 
  $Script:formt.text="Working" 
 $infoport +=  server-check 
 $pres=  $infoport| ft -AutoSize  | Out-String 
 
 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 


# endregion


# region Install Config CHG Server Function


function server-change { 

$oval =$inputbox.text
$nval =$filterbox.text
((Get-Content -path "C:\Program Files\IBM\WinCollect\config\install_config.txt" -Raw) -replace $oval,$nval) | Set-Content -Path "C:\Program Files\IBM\WinCollect\config\install_config.txt"
 
}

function CHGSER { 
progressbar 
 $outputBox.Clear() 
$statusBar.Text=("Processing the request")
 
 $date =Get-Date 
 $ct = "Task Completed @ " + $date 
 $Script:formt.Visible=$true 

  $Script:formt.text="Working on $nval" 
 $serchange +=  server-change  
 $pres=  $serchange| ft -AutoSize  | Out-String 

 $outputBox.Appendtext("{0}`n" -f $pres +"`n $ct")  
 $statusBar.Text=("Ready")
 $Script:formt.close() 
 } 


# endregion


# region Documentation Function


function SHOWDOC {
	
	
	$documentation = 
	"
 Instructions for use:
	
	WinCollect Service Buttons
	
		Check - checks the status of the WinCollect service on that computer
		Start - starts the WinCollect service on that computer
		Stop - stops the WinCollect service on that computer
		
	Port Check Buttons
		
		1) Enter the IP of the QRadar destination
		2) Enter all IPs of the systems that are being polled by or are forwarding to this server

		Port Check - with a destination server entered into the 'Enter Servers' box you can type any port number into the 'Filters' box 
		and test the connectivity (only works for TCP)
		8413 - With the IP of the QRadar managed host entered in the 'Enter Servers' box this will test the conectivity with that configuration console
		135 - With the IP of the windows systems that are forwarding to or being polled by this WinCollect server, this will test their port connectivity
		139 - With the IP of the windows systems that are forwarding to or being polled by this WinCollect server, this will test their port connectivity
		445 - With the IP of the windows systems that are forwarding to or being polled by this WinCollect server, this will test their port connectivity
	
	WinCollect Log Buttons
		
		Open - Opens the WinCollect Log file
		Search - Type a search filter e.g. IP address, or any string into the 'Filters' box and this button will return those lines in the 'Output' box
		
	Install Config Buttons
		
		Open - Opens the Install Config file
		CHK Identifier - Returns the Identifier line from the Install Config file into the 'Output' box
		CHG Identifier - Type the Current Identifier value in the 'Enter Servers' box and the new desired value in the 'Filters' box to change the Identifier
		CHK Server - Returns the Status Server line from the Install Config file into the 'Output' box
		CHG Server - Type the Current Status Server value in the 'Enter Servers' box and the new desired value in the 'Filters' box to change the Status Server 
		and Configuration Server
	
	New PEM button
	
		Generate - First click the 'Stop' button under the WinCollect Service section then press the 'Generate' button which will append '.old' to the end 
		of the PEM file. then click the 'Start' button under the WinCollect Service section which will create a new PEM file.
	"
	[System.Windows.Forms.MessageBox]::Show($documentation,"Documentation",0)
 }

# endregion


[void]$Form.ShowDialog() 
$Form.Dispose()
