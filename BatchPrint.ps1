####################################################################################################################################
#This top section is taken from https://foxdeploy.com/2015/04/16/part-ii-deploying-powershell-guis-in-minutes-using-visual-studio  #
####################################################################################################################################

#ERASE ALL THIS AND PUT XAML BELOW between the @" "@
$inputXML = @"
<Window x:Class="WpfPrintGateway.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfPrintGateway"
        mc:Ignorable="d"
        Title="Central Printing" Height="763.78" Width="1340.617" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen">
    <Grid HorizontalAlignment="Left" Height="725" Margin="10,0,0,0" VerticalAlignment="Top" Width="1179">
        <Button x:Name="buttonRefresh" Content="Refresh" HorizontalAlignment="Left" Margin="598,24,0,0" VerticalAlignment="Top" Width="75" Height="25"/>
        <Button x:Name="buttonArchive" Content="Archive" HorizontalAlignment="Left" Margin="438,24,0,0" VerticalAlignment="Top" Width="75" Height="25"/>
        <Button x:Name="buttonPrint" Content="Print" HorizontalAlignment="Left" Margin="278,24,0,0" VerticalAlignment="Top" Width="75" Height="25"/>
        <Button x:Name="buttonSendExternal" Content="Send External" HorizontalAlignment="Left" Margin="518,24,0,0" VerticalAlignment="Top" Width="75" Height="25"/>
        <ListView x:Name="listViewDocs" HorizontalAlignment="Left" Margin="0,291,0,0" Width="1169" Height="386" VerticalAlignment="Top">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Filename" Width="120" DisplayMemberBinding ="{Binding 'Filename'}"/>
                    <GridViewColumn Header="Pages" Width="45" DisplayMemberBinding ="{Binding 'Pages'}"/>
                    <GridViewColumn Header="Type" Width="100" DisplayMemberBinding ="{Binding 'Type'}"/>
                    <GridViewColumn Header="Batch Number" Width="100" DisplayMemberBinding ="{Binding 'BatchNumber'}"/>
                    <GridViewColumn Header="Ext Approved?" Width="75" DisplayMemberBinding ="{Binding 'ExtApproved'}"/>
                    <GridViewColumn Header="Doc Position" Width="75" DisplayMemberBinding ="{Binding 'DocPosition'}"/>
                        <GridViewColumn Header="Total Documents" Width="100" DisplayMemberBinding ="{Binding 'TotalDocs'}"/>
                        <GridViewColumn Header="Initiator" Width="150" DisplayMemberBinding ="{Binding 'Initiator'}"/>
                        <GridViewColumn Header="CustomerID" Width="100" DisplayMemberBinding ="{Binding 'CustomerID'}"/>
                        <GridViewColumn Header="Date" Width="120" DisplayMemberBinding ="{Binding 'Date'}"/>
                        <GridViewColumn Header="Printed" Width="50" DisplayMemberBinding ="{Binding 'Printed'}"/>
                        <GridViewColumn Header="WhenPrinted" Width="120" DisplayMemberBinding ="{Binding 'WhenPrinted'}"/>
                        <GridViewColumn Header="Folder" Width="0" DisplayMemberBinding ="{Binding 'Folder'}"/>
                </GridView>
            </ListView.View>
        </ListView>
        <ComboBox x:Name="comboBoxPrinter" HorizontalAlignment="Left" Height="25" Margin="0,24,0,0" VerticalAlignment="Top" Width="254"/>
        <TextBlock x:Name="textBlockPrinter" HorizontalAlignment="Left" Margin="0,8,0,0" TextWrapping="Wrap" Text="Printer" VerticalAlignment="Top"/>
        <TextBlock x:Name="textBlockJobs" HorizontalAlignment="Left" TextWrapping="Wrap" Text="Jobs" VerticalAlignment="Top" Margin="0,54,0,0"/>
        <TextBlock x:Name="textBlockDocuments1" HorizontalAlignment="Left" TextWrapping="Wrap" Text="Documents" VerticalAlignment="Top" Margin="0,262,0,0"/>
        <ListView x:Name="listViewJobs" HorizontalAlignment="Left" Height="182" Margin="0,75,0,0" VerticalAlignment="Top" Width="1169">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Batch Number" Width="100" DisplayMemberBinding ="{Binding 'Batch Number'}"/>
                    <GridViewColumn Header="Folder" Width="400" DisplayMemberBinding ="{Binding 'Folder'}"/>
                    <GridViewColumn Header="Route" Width="100" DisplayMemberBinding ="{Binding 'Route'}"/>
                    <GridViewColumn Header="Pre-Approved?" Width="100" DisplayMemberBinding ="{Binding 'Pre-Approved?'}"/>
                    <GridViewColumn Header="Documents" Width="100" DisplayMemberBinding ="{Binding 'Documents'}"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Button x:Name="buttonExpand" Content="Expand" HorizontalAlignment="Left" Margin="358,24,0,0" VerticalAlignment="Top" Width="75" Height="25"/>
    </Grid>
</Window>
"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*','<Window'

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try { $Form = [Windows.Markup.XamlReader]::Load($reader) }
catch { Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed." }

#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | % { Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) }

function Get-FormVariables {
  if ($global:ReadmeDisplay -ne $true) { Write-Host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow; $global:ReadmeDisplay = $true }
  #write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
  #get-variable WPF*
}


#Get-FormVariables


############################################################################
#Everything below is part of the script                                    #
############################################################################

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#===========================================================================
# Object actions
#===========================================================================

$WPFbuttonSendExternal.Add_Click({ SendExternal })
$WPFbuttonRefresh.Add_Click({ Refresh })
$WPFbuttonPrint.Add_Click({ Print })
$WPFbuttonArchive.Add_Click({ Archive })
$WPFbuttonExpand.Add_Click({ Expand })

#===========================================================================
# Store Functions
#===========================================================================

#This function will refresh the jobs that appear in the GUI
function Refresh {
$WPFlistViewJobs.Items.Clear()
$WPFlistViewDocs.Items.Clear()
$WPFbuttonPrint.Visibility = 'hidden'
  $folders = $ERPoutput | Get-ChildItem -Directory
  
  foreach ($folder in $folders) {$csvtest = $folder.fullname + "\extract.csv"
    $extractexists = Test-Path $csvtest
    if ($extractexists -eq $False) { continue }

    $documents = $folder | Get-ChildItem | where{$_.name -match "pdf"} 
    if ($documents.count -gt 1000){$route = "External"}
    Else{$route = "Internal"}

    $WPFlistviewjobs.items.Add([pscustomobject]@{'Batch Number'=$folder.name;Documents=$Documents.count;Folder=$Folder.Fullname;Route=$route})
    
  }
  $WPFcomboBoxPrinter.items.Clear()
  $Printers = Get-Printer | select Name
  foreach ($printer in $printers) { $WPFcomboBoxPrinter.AddChild($printer.Name) }
  }
  #This function will refresh the jobs that appear in the GUI

function Expand {

If ($wpflistviewjobs.selecteditems.count -gt 1){[System.Windows.Forms.MessageBox]::Show("Please only select one job")}
ElseIf($wpflistviewjobs.selecteditems.count -lt 1){[System.Windows.Forms.MessageBox]::Show("Please select a job to expand")}
Else{
    $WPFbuttonPrint.Visibility = 'visible'
    $WPFlistViewDocs.Items.Clear()

    $csvtest = $wpflistviewjobs.SelectedItem.folder + "\extract.csv"
    $extractexists = Test-Path $csvtest
    if ($extractexists -eq $False) { continue }
    $extractdetails = Import-Csv $csvtest | Select-Object @{ Name = ‘ID‘; Expression = { $_.ID } },`
       @{ Name = ‘Initiator‘; Expression = { $_.Initiator } },`
       @{ Name = ‘Type‘; Expression = { $_.Type } },`
       @{ Name = ‘Filename‘; Expression = { $_.Filename } },`
       @{ Name = ‘BatchNumber‘; Expression = { $_.batchnumber } },`
       @{ Name = ‘ExtApproved‘; Expression = { $_.externalapproved } },`
       @{ Name = ‘DocPosition‘; Expression = { $_.Docposition } },`
       @{ Name = 'TotalDocs‘; Expression = { $_.totaldocs } },`
       @{ Name = ‘Date‘; Expression = { $_.date } },`
       @{ Name = ‘Pages‘; Expression = { $_.pages } },`
       @{ Name = ‘CustomerID‘; Expression = { $_.CustomerID } },`
       @{ Name = ‘WhenPrinted‘; Expression = { $_.whenprinted } },`
       @{ Name = ‘Folder‘; Expression = { $Wpflistviewjobs.selecteditem.folder } },`
       @{ Name = ‘Printed‘; Expression = { $_.Printed } } | % { $WPFlistViewDocs.AddChild($_) }
    
                       
    }
}

#This function will print all of the jobs that have been selected in the GUI.
function Print {

$logoutput=@()
[int]$pagechecker = 0
  foreach ($line in $WPFListviewDocs.selecteditems) {
  $pages = [int]$line.pages
  $totalpages += [int]$line.pages}
  $pagechecker = $totalpages / $WPFListviewDocs.SelectedItems.count

If ($pagechecker -eq $pages){

  foreach ($item in $WPFListviewDocs.SelectedItems) {

    $PDFpath = $item.folder + "\" + $item.filename
    $SelectedPrinter = $WPFcomboBoxPrinter.SelectedValue
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = "2printer.exe"
    $pinfo.RedirectStandardError = $false
    $pinfo.RedirectStandardOutput = $false
    $pinfo.UseShellExecute = $true
    $pinfo.Arguments = "-s ""$PDFpath"" -prn ""$SelectedPrinter"" -alerts_no"
    #$pinfo.WindowStyle = 'Hidden'
    #$pinfo.CreateNoWindow = $True
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    #This section pops up the please wait form until the process has ended
    $Formwait = New-Object system.Windows.Forms.Form
    $Label = New-Object System.Windows.Forms.Label
    $formwait.Text = "Please Wait"
    $Formwait.AutoScroll = $True
    $formwait.height = 100
    $Formwait.Width = 300
    $formwait.ShowIcon = $false
    $formwait.MinimizeBox = $False
    $formwait.MaximizeBox = $False
    $formwait.WindowState = "Normal"
    $formwait.SizeGripStyle = "Hide"
    $formwait.ShowInTaskbar = $False
    $formwait.StartPosition = "CenterScreen"
    $formwait.BackColor = "White"
    $Font = New-Object System.Drawing.Font ("Tahoma",10)
    $Formwait.Font = $Font
    $Label.Text = "Printing, Please Wait."
    $Label.AutoSize = $false
    $label.textalign = "middlecenter"
    $label.Dock = "Fill"
    $formwait.Controls.Add($Label)
    $Formwait.Visible = $True
    $Formwait.Update()
    
    $p.Start() | Out-Null
    $p.StandardOutput
    #$stdout = $p.StandardOutput.ReadToEnd()
    $p.WaitForExit()
    $ExitCode = $p.ExitCode
    $JobID = $item.batchnumber

   $Formwait.Visible = $False

   if ($exitcode -eq 0) {
     $output += $item | select Filename
     $filename = $item.filename
    [System.Windows.Forms.MessageBox]::Show("Job ""$filename"" was sent to the ""$SelectedPrinter""") 
    
    $date =get-date

    $csv = $item.folder + "\extract.csv"
    $importedcsv = ""
    $importedcsv = Import-Csv $csv
    
    foreach($row in $importedcsv) 
    {
    If ($row.filename -eq $item.filename)
    {$row.printed = "Yes"
     $row.whenprinted = $date
    }

    $importedcsv | Export-Csv $csv -NoTypeInformation
    }

    }
    elseif ($Exitcode -eq 2) { [System.Windows.Forms.MessageBox]::Show("No documents found") }
    elseif ($Exitcode -eq 7) { [System.Windows.Forms.MessageBox]::Show("Printer Name is invalid") }
    else { [System.Windows.Forms.MessageBox]::Show("There was an error with the print job") }
  }

  
}

else {[System.Windows.Forms.MessageBox]::Show("Please ensure that jobs with the same number of pages have been selected")}
Expand
}

#This function checks that the folders exist and are accesible.  It also checks that the 2printer.exe application is installed.
function PreReqs {

  $ERPoutputExists = Test-Path $ERPoutput
  $ArchiveExists = Test-Path $Archive
  $sFTPExists = Test-Path $sFTP
  $2PrinterExists = Test-Path 'C:\Program Files (x86)\2printer\2printer.exe'

  if ($ERPoutputExists -eq $false) {
    [System.Windows.Forms.MessageBox]::Show("You do not have access to the ERPoutput folder")
    exit
  }
  elseif ($ArchiveExists -eq $false) {
    [System.Windows.Forms.MessageBox]::Show("You do not have access to the Archive folder")
    exit
  }
  elseif ($sFTPExists -eq $false) {
    [System.Windows.Forms.MessageBox]::Show("You do not have access to the sFTP folder")
    exit
  }
  elseif ($2PrinterExists -eq $false) {
    [System.Windows.Forms.MessageBox]::Show("2printer is not installed")
    exit
  }

}

#This function will move the selected jobs in the GUI into the archive folder for deletion
function Archive {

If ($wpflistviewjobs.selecteditems.count -lt 1){[System.Windows.Forms.MessageBox]::Show("Please select a job to archive")}
Else{
  foreach ($item in $WPFListviewJobs.SelectedItems) {
    $JobID = $item.'batch number'
    $Answer = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to Archive $JobID","Status",4)
    if ($Answer -eq "Yes") { Move-Item -Path $item.folder -Destination $Archive }
    elseif ($Answer -eq "No") {}
  }
  Refresh
  }
}

#This function will move the selected items to a different folder for sFTP transfer
function SendExternal {
  foreach ($item in $WPFListviewjobs.SelectedItems) {
    $JobID = $item.'batch number'
    $Answer = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to Send $JobID for external printing","Status",4)
    if ($Answer -eq "Yes") { Move-Item -Path $item.folder -Destination $sFTP }
    elseif ($Answer -eq "No") {}
  }
  Refresh
}

#===========================================================================
# Shows the form
#===========================================================================

$ERPOutput = "C:\PrintGateway\TeamasBatchPrint\ERPOutput"
$Archive = "c:\PrintGateway\TeamasBatchPrint\Archive"
$sFTP = "c:\PrintGateway\TeamasBatchPrint\sFTP"
PreReqs
Refresh

$Form.ShowDialog()
