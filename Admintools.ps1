### Add a new Menu Item:
### 1) Create a new MenuItem  (System.Windows.Forms.MenuItem)
### 2) Add an action to the MenuItem ($Example.Add_Click)
##### for cmd --> Start-Process -FilePath example.cmd
##### for msc --> Start-Process -FilePath "C:\Windows\System32\mmc.exe" -ArgumentList "example.msc" -Verb RunAs
### 3) Add the new MenuItem to the Menu: $Admintool.ContextMenu.MenuItems.AddRange($Example) !!Alphabetical order!!
### 4) Test if the action can be executed 

#Assemblies 
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')    | out-null
[System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework')   | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')          | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | out-null
 
#Icon Systray Tool
$Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\mmc.exe")
 
#Systray Object 
$Admintool = New-Object System.Windows.Forms.NotifyIcon
$Admintool.Text = "Admintools"
$Admintool.Icon = $Icon
$Admintool.Visible = $true
 
#Menu Certificate
$MenuCert = New-Object System.Windows.Forms.MenuItem
$MenuCert.Text = "Certificate"
$MenuCert.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\mmc.exe" -ArgumentList "certmgr.msc" -Verb RunAs
}) 

#Menu Computermanagement
$MenuCompMgmt = New-Object System.Windows.Forms.MenuItem
$MenuCompMgmt.Text = "Computer Management"
$MenuCompMgmt.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\mmc.exe" -ArgumentList "compmgmt.msc" -Verb RunAs
})

#Menu Date and Time
$MenuTimedate = New-Object System.Windows.Forms.MenuItem
$MenuTimedate.Text = "Date and Time"
$MenuTimedate.Add_Click({ 
    $PathTimeDate = "C:\Program Files (x86)\Ammann\AdminTools\Bin\AdminTimedate.cmd"
    $CheckTimeDate = Test-Path $PathTimeDate
    if($CheckTimedate -eq $true){
        Start-Process -FilePath $PathTimeDate
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathTimeDate", 'Path not found', 'Ok', 'Error') 
    }   
})

#Menu Device Manager
$MenuDevMgmt = New-Object System.Windows.Forms.MenuItem
$MenuDevMgmt.Text = "Device Manager"
$MenuDevMgmt.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\mmc.exe" -ArgumentList "devmgmt.msc" -Verb RunAs
}) 

#Menu File Manager
$MenuFreeComm = New-Object System.Windows.Forms.MenuItem
$MenuFreeComm.Text = "File Manager (FreeCommander)"
$MenuFreeComm.Add_Click({ 
    $PathFreeComm = "C:\Program Files (x86)\Ammann\AdminTools\Bin\FreeCommander.cmd"
    $CheckFreeComm = Test-Path $PathFreeComm
    if($CheckFreeComm -eq $true){
        Start-Process -FilePath $PathFreeComm
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathFreeComm", 'Path not found', 'Ok', 'Error') 
    } 
}) 

#Menu Group Policy
$MenuGrpPol = New-Object System.Windows.Forms.MenuItem
$MenuGrpPol.Text = "Group Policy"
$MenuGrpPol.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\mmc.exe" -ArgumentList "gpedit.msc" -Verb RunAs
})  

#Menu IIS Manager
$MenuIISMgr = New-Object System.Windows.Forms.MenuItem
$MenuIISMgr.Text = "IIS Manager"
$MenuIISMgr.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\inetsrv\InetMgr.exe" -Verb RunAs
})  

#Menu Local Security Policy
$MenuSecPol = New-Object System.Windows.Forms.MenuItem
$MenuSecPol.Text = "Local Security Policy"
$MenuSecPol.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\mmc.exe" -ArgumentList "secpol.msc" -Verb RunAs
})  

#Menu Local Users and Groups
$MenuLUsrMgr = New-Object System.Windows.Forms.MenuItem
$MenuLUsrMgr.Text = "Local Users and Groups"
$MenuLUsrMgr.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\mmc.exe" -ArgumentList "lusrmgr.msc" -Verb RunAs
}) 

#Menu Network Connections
$MenuNetCon = New-Object System.Windows.Forms.MenuItem
$MenuNetCon.Text = "Network Connections"
$MenuNetCon.Add_Click({ 
    Start-Process -FilePath "C:\Windows\explorer.exe" -ArgumentList "/e,::{7007acc7-3202-11d1-aad2-00805fc1270e}" -Verb RunAs
}) 

#Menu Printer
$MenuPrinter = New-Object System.Windows.Forms.MenuItem
$MenuPrinter.Text = "Printer"
$MenuPrinter.Add_Click({ 
    Start-Process -FilePath "C:\Windows\explorer.exe" -ArgumentList "/e,::{2227A280-3AEA-1069-A2DE-08002B30309D}" -Verb RunAs
}) 

#Menu Prompt cmd (ComaSystem)
$MenuAdminPromptCmd = New-Object System.Windows.Forms.MenuItem
$MenuAdminPromptCmd.Text = "Admin Prompt CMD (ComaSystem)"
$MenuAdminPromptCmd.Add_Click({ 
    $PathAdminPromptCmd = "C:\Program Files (x86)\Ammann\AdminTools\Bin\adminprompt.cmd"
    $CheckAdminPromptCmd = Test-Path $PathAdminPromptCmd
    if($CheckAdminPromptCmd -eq $true){
        Start-Process -FilePath $PathAdminPromptCmd
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathAdminPromptCmd", 'Path not found', 'Ok', 'Error') 
    }
}) 

#Menu Prompt PowerShell
$MenuAdminPromptPS = New-Object System.Windows.Forms.MenuItem
$MenuAdminPromptPS.Text = "Admin Prompt PowerShell (ComaSystem)"
$MenuAdminPromptPS.Add_Click({ 
    $PathAdminPromptPS = "C:\Program Files (x86)\Ammann\AdminTools\Bin\adminpromptps.cmd"
    $CheckAdminPromptPS = Test-Path $PathAdminPromptPS
    if($CheckAdminPromptPS -eq $true){
        Start-Process -FilePath $PathAdminPromptPS
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathAdminPromptPS", 'Path not found', 'Ok', 'Error') 
    }
}) 

#Menu Registry Editor
$MenuRegEdit = New-Object System.Windows.Forms.MenuItem
$MenuRegEdit.Text = "Registry Editor"
$MenuRegEdit.Add_Click({ 
    $PathRegEdit = "C:\Program Files (x86)\Ammann\AdminTools\Bin\RegEdit.cmd"
    $CheckRegEdit = Test-Path $PathRegEdit
    if($CheckRegEdit -eq $true){
        Start-Process -FilePath $PathRegEdit
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathRegEdit", 'Path not found', 'Ok', 'Error') 
    }
}) 

#Menu Services
$MenuServices = New-Object System.Windows.Forms.MenuItem
$MenuServices.Text = "Services"
$MenuServices.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\mmc.exe" -ArgumentList "services.msc" -Verb RunAs
})

#Menu Software
$MenuSoftware = New-Object System.Windows.Forms.MenuItem
$MenuSoftware.Text = "Software"
$MenuSoftware.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\rundll32.exe" -ArgumentList "shell32.dll,Control_RunDLL appwiz.cpl" -Verb RunAs
}) 

#Menu SQL Server Management Studio 16
$MenuSQLMgt16 = New-Object System.Windows.Forms.MenuItem
$MenuSQLMgt16.Text = "SQL Server Management Studio 16"
$MenuSQLMgt16.Add_Click({ 
    $PathSQLMgt16 = "C:\Program Files (x86)\Microsoft SQL Server\130\Tools\Binn\ManagementStudio\Ssms.exe"
    $CheckSQLMgt16 = Test-Path $PathSQLMgt16
    if($CheckSQLMgt16 -eq $true){
        Start-Process -FilePath $PathSQLMgt16 -Verb RunAs
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathSQLMgt16", 'Path not found', 'Ok', 'Error') 
    }
}) 

#Menu SQL Server Management Studio 17
$MenuSQLMgt17 = New-Object System.Windows.Forms.MenuItem
$MenuSQLMgt17.Text = "SQL Server Management Studio 17"
$MenuSQLMgt17.Add_Click({ 
    $PathSQLMgt17 = "C:\Program Files (x86)\Microsoft SQL Server\140\Tools\Binn\ManagementStudio\Ssms.exe"
    $CheckSQLMgt17 = Test-Path $PathSQLMgt17
    if($CheckSQLMgt17 -eq $true){
        Start-Process -FilePath $PathSQLMgt17 -Verb RunAs
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathSQLMgt17", 'Path not found', 'Ok', 'Error') 
    }
})

#Menu System Properties
$MenuSysProp = New-Object System.Windows.Forms.MenuItem
$MenuSysProp.Text = "System Properties"
$MenuSysProp.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\rundll32.exe" -ArgumentList "shell32.dll,Control_RunDLL sysdm.cpl" -Verb RunAs
})

#Menu Task Scheduler
$MenuTaskSched = New-Object System.Windows.Forms.MenuItem
$MenuTaskSched.Text = "Task Scheduler"
$MenuTaskSched.Add_Click({ 
    Start-Process -FilePath "C:\Windows\System32\taskschd.msc" -Verb RunAs 
}) 

#Menu Taskmanger
$MenuTaskMgr = New-Object System.Windows.Forms.MenuItem
$MenuTaskMgr.Text = "Taskmanager"
$MenuTaskMgr.Add_Click({ 
    $PathTaskMgr = "C:\Windows\System32\taskmgr.exe"
    $CheckTaskMgr = Test-Path $PathTaskMgr
    if($CheckTaskMgr -eq $true){
        Start-Process -FilePath $PathTaskMgr -Verb RunAs
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathTaskMgr", 'Path not found', 'Ok', 'Error') 
    }
})  

#Menu Reset ProfiNet
$MenuResProfinet = New-Object System.Windows.Forms.MenuItem
$MenuResProfinet.Text = "Reset Profinet"
$MenuResProfinet.Add_Click({ 
    $PathResProfinet = "C:\Program Files (x86)\Ammann\AdminTools\Bin\reset_profinet.bat"
    $CheckResProfinet = Test-Path $PathResProfinet
    if($CheckResProfinet -eq $true){
        Start-Process -FilePath $PathResProfinet -Verb RunAs
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathResProfinet", 'Path not found', 'Ok', 'Error') 
    }
})

#Menu as1DisasterRecovery
$MenuDisRec = New-Object System.Windows.Forms.MenuItem
$MenuDisRec.Text = "as1DisasterRecovery"
$MenuDisRec.Add_Click({ 
    $PathDisRec = "C:\Program Files (x86)\Ammann\as1DisasterRecovery\as1DisasterRecovery.exe"
    $CheckDisRec = Test-Path $PathDisRec
    if($CheckDisRec -eq $true){
        Start-Process -FilePath $PathDisRec -Verb RunAs
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathDisRec", 'Path not found', 'Ok', 'Error') 
    }
})

#Menu NetID
$MenuNetID = New-Object System.Windows.Forms.MenuItem
$MenuNetID.Text = "NetID"
$MenuNetID.Add_Click({ 
    $PathNetID = "C:\Program Files (x86)\Ammann\AdminTools\Bin\NetID.cmd"
    $CheckNetID = Test-Path $PathNetID
    if($CheckNetID -eq $true){
        Start-Process -FilePath $PathNetID
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathNetID", 'Path not found', 'Ok', 'Error') 
    }
})

#Menu Proxy_Configuration
$MenuProxyConf = New-Object System.Windows.Forms.MenuItem
$MenuProxyConf.Text = "Proxy Configuration"
$MenuProxyConf.Add_Click({ 
    $PathProxyConf = "C:\Program Files (x86)\Ammann\AdminTools\Bin\proxy_gui_starter.cmd"
    $CheckProxyConf = Test-Path $PathProxyConf
    if($CheckProxyConf -eq $true){
        Start-Process -FilePath $PathProxyConf
    } else {
        [System.Windows.MessageBox]::Show("Couldn't find $PathProxyConf", 'Path not found', 'Ok', 'Error') 
    }
})

#Context Menu  /// Alphabetical order !!
$ContextMenu = New-Object System.Windows.Forms.ContextMenu
$Admintool.ContextMenu = $ContextMenu
$Admintool.ContextMenu.MenuItems.AddRange($MenuAdminPromptCmd)
$Admintool.ContextMenu.MenuItems.AddRange($MenuAdminPromptPS)
$Admintool.ContextMenu.MenuItems.AddRange($MenuDisRec)
$Admintool.ContextMenu.MenuItems.AddRange($MenuCert)
$Admintool.ContextMenu.MenuItems.AddRange($MenuCompMgmt)
$Admintool.ContextMenu.MenuItems.AddRange($MenuTimedate)
$Admintool.ContextMenu.MenuItems.AddRange($MenuDevMgmt)
$Admintool.ContextMenu.MenuItems.AddRange($MenuFreeComm)
$Admintool.ContextMenu.MenuItems.AddRange($MenuGrpPol)
$Admintool.ContextMenu.MenuItems.AddRange($MenuIISMgr)
$Admintool.ContextMenu.MenuItems.AddRange($MenuSecPol)
$Admintool.ContextMenu.MenuItems.AddRange($MenuLUsrMgr)
$Admintool.ContextMenu.MenuItems.AddRange($MenuNetID)
$Admintool.ContextMenu.MenuItems.AddRange($MenuNetCon)
$Admintool.ContextMenu.MenuItems.AddRange($MenuPrinter)
$Admintool.ContextMenu.MenuItems.AddRange($MenuProxyConf)
$Admintool.ContextMenu.MenuItems.AddRange($MenuRegEdit)
$Admintool.ContextMenu.MenuItems.AddRange($MenuResProfinet)
$Admintool.ContextMenu.MenuItems.AddRange($MenuSQLMgt16)
$Admintool.ContextMenu.MenuItems.AddRange($MenuSQLMgt17)
$Admintool.ContextMenu.MenuItems.AddRange($MenuServices)
$Admintool.ContextMenu.MenuItems.AddRange($MenuSoftware)
$Admintool.ContextMenu.MenuItems.AddRange($MenuSysProp)
$Admintool.ContextMenu.MenuItems.AddRange($MenuTaskSched)
$Admintool.ContextMenu.MenuItems.AddRange($MenuTaskMgr)

#let PowerShell disappear
$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

#Garbage Collector to lower RAM usage
[System.GC]::Collect()
 
#Application context to run all within
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)
