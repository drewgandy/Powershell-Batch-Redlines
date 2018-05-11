<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Batch Redlines
#>
Set-ExecutionPolicy unrestricted
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$FrmMain                         = New-Object system.Windows.Forms.Form
$FrmMain.ClientSize              = '800,500'
$FrmMain.text                    = "Batch Redlines"
$FrmMain.TopMost                 = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Original"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(50,10)
$Label1.Font                     = 'Arial,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Modified"
$Label2.AutoSize                 = $true
$Label2.width                    = 42
$Label2.height                   = 10
$Label2.Anchor                   = 'top,right'
$Label2.location                 = New-Object System.Drawing.Point(675,10)
$Label2.Font                     = 'Arial,10'

$LstOriginal                     = New-Object system.Windows.Forms.ListBox
$LstOriginal.text                = "listBox"
$LstOriginal.width               = 345
$LstOriginal.height              = 275
$LstOriginal.location            = New-Object System.Drawing.Point(50,40)
$LstOriginal.IntegralHeight      = $True
$LstOriginal.HorizontalScrollbar = $True
$LstOriginal.AllowDrop           = $True

$LstModified                     = New-Object system.Windows.Forms.ListBox
$LstModified.text                = "listBox"
$LstModified.width               = 345
$LstModified.height              = 275
$LstModified.Anchor              = 'top,right'
$LstModified.location            = New-Object System.Drawing.Point(405,40)
$LstModified.IntegralHeight      = $True
$LstModified.HorizontalScrollbar = $True
$LstModified.AllowDrop           = $True

$CmdOriginalMoveUp               = New-Object system.Windows.Forms.Button
$CmdOriginalMoveUp.text          = "˄"
$CmdOriginalMoveUp.width         = 30
$CmdOriginalMoveUp.height        = 30
$CmdOriginalMoveUp.location      = New-Object System.Drawing.Point(10,40)
$CmdOriginalMoveUp.Font          = 'Arial,10'

$CmdModifiedMoveUp               = New-Object system.Windows.Forms.Button
$CmdModifiedMoveUp.text          = "˄"
$CmdModifiedMoveUp.width         = 30
$CmdModifiedMoveUp.height        = 30
$CmdModifiedMoveUp.location      = New-Object System.Drawing.Point(760,40)
$CmdModifiedMoveUp.Anchor        = 'top,right'
$CmdModifiedMoveUp.Font          = 'Arial,10'

$CmdOriginalMoveDown             = New-Object system.Windows.Forms.Button
$CmdOriginalMoveDown.text        = "˅"
$CmdOriginalMoveDown.width       = 30
$CmdOriginalMoveDown.height      = 30
$CmdOriginalMoveDown.location    = New-Object System.Drawing.Point(10,275)
$CmdOriginalMoveDown.Font        = 'Arial,10'

$CmdModifiedMoveDown             = New-Object system.Windows.Forms.Button
$CmdModifiedMoveDown.text        = "˅"
$CmdModifiedMoveDown.width       = 30
$CmdModifiedMoveDown.height      = 30
$CmdModifiedMoveDown.location    = New-Object System.Drawing.Point(760,275)
$CmdModifiedMoveDown.Anchor        = 'top,right'
$CmdModifiedMoveDown.Font        = 'Arial,10'

$CmdOriginalDelete               = New-Object system.Windows.Forms.Button
$CmdOriginalDelete.text          = "Delete Selected"
$CmdOriginalDelete.width         = 175
$CmdOriginalDelete.height        = 30
$CmdOriginalDelete.location      = New-Object System.Drawing.Point(50,310)
$CmdOriginalDelete.Font          = 'Arial,10'

$CmdModifiedDelete               = New-Object system.Windows.Forms.Button
$CmdModifiedDelete.text          = "Delete Selected"
$CmdModifiedDelete.width         = 175
$CmdModifiedDelete.height        = 30
$CmdModifiedDelete.location      = New-Object System.Drawing.Point(575,310)
$CmdModifiedDelete.Font          = 'Arial,10'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Output Folder:"
$Label3.AutoSize                 = $true
$Label3.width                    = 190
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(10,360)
$Label3.Font                     = 'Arial,10'

$TxtOutputFolder                 = New-Object system.Windows.Forms.TextBox
$TxtOutputFolder.multiline       = $false
$TxtOutputFolder.width           = 337
$TxtOutputFolder.height          = 20
$TxtOutputFolder.location        = New-Object System.Drawing.Point(105,357)
$TxtOutputFolder.Font            = 'Arial,10'
$TxtOutputFolder.text            = [environment]::getfolderpath("UserProfile") +"\Downloads" #"C:\Users\dhgandy\Downloads"

$CmdOutputFolderBrowse           = New-Object system.Windows.Forms.Button
$CmdOutputFolderBrowse.text      = ". . ."
$CmdOutputFolderBrowse.width     = 60
$CmdOutputFolderBrowse.height    = 30
$CmdOutputFolderBrowse.location  = New-Object System.Drawing.Point(445,355)
$CmdOutputFolderBrowse.Font      = 'Arial,10'

$TxtRedlineName                  = New-Object system.Windows.Forms.TextBox
$TxtRedlineName.multiline        = $false
$TxtRedlineName.width            = 100
$TxtRedlineName.height           = 20
$TxtRedlineName.location         = New-Object System.Drawing.Point(105,390)
$TxtRedlineName.Font             = 'Arial,10'
$TxtRedlineName.text             = "REDLINE - "
$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Redline Name:"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(10,393)
$Label4.Font                     = 'Arial,10'

$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "[DOCUMENT NAME] ."
$Label5.AutoSize                 = $true
$Label5.width                    = 25
$Label5.height                   = 10
$Label5.location                 = New-Object System.Drawing.Point(210,393)
$Label5.Font                     = 'Arial,10'

$CmbRedlineFileType              = New-Object system.Windows.Forms.ComboBox
$CmbRedlineFileType.text         = "PDF"
$CmbRedlineFileType.width        = 85
$CmbRedlineFileType.height       = 20
@('PDF', 'DOC', 'RTF') | ForEach-Object {[void] $CmbRedlineFileType.Items.Add($_)}
$CmbRedlineFileType.location     = New-Object System.Drawing.Point(355,391)
$CmbRedlineFileType.Font         = 'Arial,10'

$CmdRunRedlines                  = New-Object system.Windows.Forms.Button
$CmdRunRedlines.text             = "Run Redlines"
$CmdRunRedlines.width            = 140
$CmdRunRedlines.height           = 30
$CmdRunRedlines.location         = New-Object System.Drawing.Point(655,450)
$CmdRunRedlines.Font             = 'Arial,10'

$FrmMain.controls.AddRange(@($Label1,$Label2,$LstOriginal,$LstModified,$CmdOriginalMoveUp,$CmdModifiedMoveUp,$CmdOriginalMoveDown,$CmdModifiedMoveDown,$CmdOriginalDelete,$CmdModifiedDelete,$Label3,$TxtOutputFolder,$CmdOutputFolderBrowse,$TxtRedlineName,$Label4,$Label5,$CmbRedlineFileType,$CmdRunRedlines))

#region gui events {



$LstOriginal_DragDrop = [System.Windows.Forms.DragEventHandler]{
	foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
	{
		$LstOriginal.Items.Add($filename)
	}
}

$LstOriginal_DragOver = [System.Windows.Forms.DragEventHandler]{
	if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
	{
	    $_.Effect = 'Copy'
	}
	else
	{
	    $_.Effect = 'None'
	}
}
$LstModified_DragOver = [System.Windows.Forms.DragEventHandler]{
	if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
	{
	    $_.Effect = 'Copy'
	}
	else
	{
	    $_.Effect = 'None'
	}
}

$LstModified_DragDrop = [System.Windows.Forms.DragEventHandler]{
	foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
	{
		$LstModified.Items.Add($filename)
	}
}


$CmdModifiedMoveDown_Click=
{
#only if the last item isn't the current one
   if(($LstModified.SelectedIndex -ne -1)   -and   ($LstModified.SelectedIndex -lt $LstModified.Items.Count - 1)    )   {
        $LstModified.BeginUpdate()
        #Get starting position
        $pos = $LstModified.selectedIndex
        # add a duplicate of item below in the listbox
        $LstModified.items.insert($pos,$LstModified.Items.Item($pos +1))
        # delete the old occurrence of this item
        $LstModified.Items.RemoveAt($pos +2 )
        # move to current item
        $LstModified.SelectedIndex = ($pos +1)
        $LstModified.EndUpdate()
   }ELSE{
       #Bottom of list, beep
       [console]::beep(500,100)
   }
}

$CmdModifiedMoveUp_Click=
{
    if($LstModified.SelectedIndex -gt 0)
    {
        $LstModified.BeginUpdate()
        #Get starting position
        $pos = $LstModified.selectedIndex
        # add a duplicate of original item up in the listbox
        $LstModified.items.insert($pos -1,$LstModified.Items.Item($pos))
        # make it the current item
        $LstModified.SelectedIndex = ($pos -1)
        # delete the old occurrence of this item
        $LstModified.Items.RemoveAt($pos +1)
        $LstModified.EndUpdate()
    }ELSE{
       #Top of list, beep
       [console]::beep(500,100)
   }
}

$CmdOriginalMoveDown_Click=
{
#only if the last item isn't the current one
   if(($LstOriginal.SelectedIndex -ne -1)   -and   ($LstOriginal.SelectedIndex -lt $LstOriginal.Items.Count - 1)    )   {
        $LstOriginal.BeginUpdate()
        #Get starting position
        $pos = $LstOriginal.selectedIndex
        # add a duplicate of item below in the listbox
        $LstOriginal.items.insert($pos,$LstOriginal.Items.Item($pos +1))
        # delete the old occurrence of this item
        $LstOriginal.Items.RemoveAt($pos +2 )
        # move to current item
        $LstOriginal.SelectedIndex = ($pos +1)
        $LstOriginal.EndUpdate()
   }ELSE{
       #Bottom of list, beep
       [console]::beep(500,100)
   }
}

$CmdOriginalMoveUp_Click=
{
    if($LstOriginal.SelectedIndex -gt 0)
    {
        $LstOriginal.BeginUpdate()
        #Get starting position
        $pos = $LstOriginal.selectedIndex
        # add a duplicate of original item up in the listbox
        $LstOriginal.items.insert($pos -1,$LstOriginal.Items.Item($pos))
        # make it the current item
        $LstOriginal.SelectedIndex = ($pos -1)
        # delete the old occurrence of this item
        $LstOriginal.Items.RemoveAt($pos +1)
        $LstOriginal.EndUpdate()
    }ELSE{
       #Top of list, beep
       [console]::beep(500,100)
   }
}

$CmdOutputFolderBrowse_Click=
{
	
	Add-Type -AssemblyName System.Windows.Forms
	$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
	[void]$FolderBrowser.ShowDialog()
	$TxtOutputFolder.text = $FolderBrowser.SelectedPath
	
}

$CmdModifiedDelete_Click=
{
    if ($LstModified.SelectedIndex -ge 0)
		{
        $i = $LstModified.SelectedIndex
        $LstModified.Items.RemoveAt($LstModified.SelectedIndex)
        if ($i -le $LstModified.Items.Count - 1)
            {
            $LstModified.SetSelected($i, $True)
            }ELSE
            {
            if($LstModified.Items.Count -ne 0)
            {$LstModified.SetSelected($LstModified.items.Count-1, $True)}
            }

        }

}


$CmdOriginalDelete_Click=
{
	if ($LstOriginal.SelectedIndex -ge 0) 
        {$i = $LstOriginal.SelectedIndex
		$LstOriginal.Items.RemoveAt($LstOriginal.SelectedIndex)
        if ($i -le $LstOriginal.Items.Count - 1)
            {
            $LstOriginal.SetSelected($i, $True)
            }ELSE
            {
            if($LstOriginal.Items.Count -ne 0)
            {$LstOriginal.SetSelected($LstOriginal.items.Count-1, $True)}
            }
        }
}

$CmdRunRedlines_Click=
{
	if ($LstOriginal.Items.Count -ne $LstModified.Items.Count) 
		{[System.Windows.MessageBox]::Show('The number of modified documents does not match the number of original documents.  Please ensure there are corresponding documents between the two lists.')
    }ELSE{
        if ($txtOutputFolder.text -eq "")
            {
        	Add-Type -AssemblyName System.Windows.Forms
        	$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
	        [void]$FolderBrowser.ShowDialog()
	        $TxtOutputFolder.text = $FolderBrowser.SelectedPath
            }
        #if ($txtOutputFolder.text -eq ""){Return}
        for ($i=0; $i -lt $LstOriginal.Items.Count; $i++)
			{
                $LstOriginal.SetSelected($i, $True)
                $LstModified.SetSelected($i, $True)

                $OriginalFilename = '/original="' + $LstOriginal.Items[$i] + '"'
                Write-Host '/v ' + $OriginalFilename
                $ModifiedFilename = '/modified="' + $LstModified.Items[$i] + '"'
                Write-Host $ModifiedFilename
                $Outputfilename = '/outfile="' + $txtOutputFolder.text.trim('\') + '\' + $txtRedlineName.text + [io.path]::GetFileNameWithoutExtension($LstModified.Items[$i]) + '.' + $CmbRedlineFileType.text  + '"'
                Write-Host $OutputFilename
                Start-Process -FilePath "deltavw.exe" -ArgumentList '/v', $OriginalFilename, $Modifiedfilename, $Outputfilename -Wait 
 			}
        [System.Windows.MessageBox]::Show('Finished running redlines.')
 
    }
}

$LstOriginal_Click=
{
    if(($LstModified.Items.Count -ge $LstOriginal.SelectedIndex) -and ($LstModified.Items.Count -ne 0))
        {
            $LstModified.SetSelected($LstOriginal.SelectedIndex, $True)
        }
}

$LstModified_Click=
{ 
    if(($LstOriginal.Items.Count -ge $LstModified.SelectedIndex) -and ($LstOriginal.Items.Count -ne 0))
        {
            $LstOriginal.SetSelected($LstModified.SelectedIndex, $True)
        }

}
#endregion events }
$LstOriginal.Add_DragDrop($LstOriginal_DragDrop)
$LstOriginal.Add_DragOver($LstOriginal_DragOver)
$LstOriginal.Add_Click($LstOriginal_Click)
$LstModified.Add_DragDrop($LstModified_DragDrop)
$LstModified.Add_DragOver($LstModified_DragOver)
$LstModified.Add_Click($LstModified_Click)
$CmdModifiedMoveUp.Add_Click($CmdModifiedMoveUp_Click)
$CmdModifiedMoveDown.Add_Click($CmdModifiedMoveDown_Click)
$CmdOriginalMoveUp.Add_Click($CmdOriginalMoveUp_Click)
$CmdOriginalMoveDown.Add_Click($CmdOriginalMoveDown_Click) 
$CmdOutputFolderBrowse.Add_Click($CmdOutputFolderBrowse_Click)
$CmdRunRedlines.Add_Click($CmdRunRedlines_Click)
$CmdModifiedDelete.Add_Click($CmdModifiedDelete_Click)
$CmdOriginalDelete.Add_Click($CmdOriginalDelete_Click)

#endregion GUI }


#Write your logic code here

[void]$FrmMain.ShowDialog()