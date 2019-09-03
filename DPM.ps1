function Push-Scrip{
  param (
    [Parameter(Mandatory=$true)]$Computer,
    [Parameter(Mandatory=$true)]$RemoteComputerFolder,
    [Parameter(Mandatory=$true)]$processNumberMax,
    [Parameter(Mandatory=$true)]$OutPutFile
  )
  
  $activeProcesses = 1
  #Starts a 1 second pause until we have less Jobs than the max number
  do {
    
    Start-Sleep -Seconds 1
        
    #Remove Completed Jobs
    foreach ($job in Get-Job) {
      if ( $job.State -eq "Completed" ) { Remove-Job -Id $job.Id }
    }
        
    #Update number of active processes
    $activeProcesses = (Get-Job).Count

  }until ($activeProcesses -le $processNumberMax)


  Start-Job -ScriptBlock {

    $Computer = $args[0]
    $OutPutFile = $args[1]
    $RemoteComputerFolder = $args[2]

    $netCon = Test-NetConnection -ComputerName $Computer -CommonTCPPort WINRM
    

    if ($netCon.TcpTestSucceeded){

      Invoke-Command -ScriptBlock{

        $RemoteComputerFolder = $args[0]
                
        $getDefP = $RemoteComputerFolder+"\getDefP.ps1"
        $runner = $RemoteComputerFolder+"\runner.bat"

        ##############  Remote Scrit Definition Start:
        $Line01 = "Set-Variable -Name DefaultPrinter -Scope Global -Force`r`n"

        $Line02 = "If (Test-Path `""+$RemoteComputerFolder+"\DefaultPrinter.txt`") {`r`n"
        $Line03 = "    Remove-Item -Path `""+$RemoteComputerFolder+"\DefaultPrinter.txt`" -Force`r`n"
        $Line04 = "}`r`n"
        $Line05 = "`$DefaultPrinter = Get-WmiObject -Class win32_printer -ComputerName `"localhost`" -Filter `"Default='true'`" | Select-Object ShareName`r`n"
                
        $Line06 = "If (`$DefaultPrinter.ShareName -ne `$null) {`r`n"
        $Line07 = "    `$DefaultPrinter.ShareName | Out-File -FilePath `""+$RemoteComputerFolder+"\DefaultPrinter.txt`" -Force -Encoding `"ASCII`"`r`n"
        $Line08 = "}else{`r`n"
        $Line09 = "   `$DefaultPrinter = `"No Default Printer`"`r`n"
        $Line10= "    `$DefaultPrinter | Out-File -FilePath `""+$RemoteComputerFolder+"\DefaultPrinter.txt`" -Force -Encoding `"ASCII`"`r`n"
        $Line11 = "}`r`n"
                
        $Line12 = "#Cleanup Global Variables`r`n"
        $Line13 = "Remove-Variable -Name DefaultPrinter -Scope Global -Force`r`n"
        [string]$GetPrinter = $Line01,$Line02,$Line03,$Line04,$Line05,$Line06,$Line07,$Line08,$Line09,$Line10,$Line11,$Line12,$Line13
        ##############  Remote Scrit Definition Ends
        [string]$RunnerScript = "PowerShell.exe -Command `"& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -InputFormat None -ExecutionPolicy Bypass -File `"`"c:\2r23edgtrg67u8iyhjrtegw231refqw2356yh\getDefP.ps1`"`"' -Verb RunAs}`"`r`n"

        if (Test-Path -Path ($RemoteComputerFolder)){Remove-Item -Path $RemoteComputerFolder -Recurse -Force}
        New-Item -Path $remoteComputerFolder -ItemType Directory

        foreach ($Line in $GetPrinter){Out-File -FilePath $getDefP -InputObject $Line -Encoding ascii}
        Out-File -FilePath $runner -InputObject $RunnerScript -Encoding ascii

        #c:\ProgramData
        $strAllUsersProfile = [io.path]::GetFullPath($env:AllUsersProfile)
        $objShell = New-Object -com "Wscript.Shell"
        $linkPath = $strAllUsersProfile + "\Start Menu\Programs\Startup\getPrn.lnk"
        if (Test-Path -Path ($linkPath)){Remove-Item -Path $linkPath -Recurse -Force}
        $objShortcut = $objShell.CreateShortcut($strAllUsersProfile + "\Start Menu\Programs\Startup\getPrn.lnk")
        $objShortcut.TargetPath = $runner
        $objShortcut.Save()
        

      } -ComputerName $Computer -ArgumentList $RemoteComputerFolder

      $l = $RemoteComputerFolder.Length - 3
      $RemoteComputerFolder = $RemoteComputerFolder.Substring(3,$l) #Remove c:\ from the string

      $getDefP = "\\"+$Computer+"\c$\"+$RemoteComputerFolder+"\getDefP.ps1"
      $runner = "\\"+$Computer+"\c$\"+$RemoteComputerFolder+"\runner.bat"
      $RemoteLink = "\\"+$Computer+"\c$\ProgramData\Start Menu\Programs\Startup\getPrn.lnk"

            if( (Test-Path -Path $getDefP) -and (Test-Path -Path $runner) -and (Test-Path -Path $RemoteLink) ) {
              $OutString = $Computer+",Push Succeded" 
              Out-File -Append -FilePath $OutPutFile -InputObject $OutString
              $Date = (Get-Date).AddMinutes(3)

              $message = "For administrative reasons we need you to restart the computer. If the computer is not restarted before "+$Date+", it will restart automatically. Your IT Team."
              Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $message" -ComputerName $Computer

              $NetComputer = "\\"+$Computer
              #To change the time replace the number after /t for the number of SECONDS to wait until the system is shutdown
              SHUTDOWN /r /f /t 600 /m $NetComputer /c $message
              
            }else{
              Out-File -Append -FilePath $OutPutFile -InputObject $Computer
            }
        }else{
            Out-File -Append -FilePath $OutPutFile -InputObject $Computer
        }

    } -ArgumentList $Computer, $OutPutFile, $RemoteComputerFolder, $OutPutFile > $null

}

function Get-Data{
  param (
    [Parameter(Mandatory=$true)]$Computer,
    [Parameter(Mandatory=$true)]$RemoteComputerFolder,
    [Parameter(Mandatory=$true)]$processNumberMax,
    [Parameter(Mandatory=$true)]$OutPutFile
  )
  
  $activeProcesses = 1
  #Starts a 1 second pause until we have less Jobs than the max number
  do {
    
      Start-Sleep -Seconds 1
        
      #Remove Completed Jobs
      foreach ($job in Get-Job) {
          if ( $job.State -eq "Completed" ) { Remove-Job -Id $job.Id }
      }
        
      #Update number of active processes
      $activeProcesses = (Get-Job).Count

  }until ($activeProcesses -le $processNumberMax)
  
  
  Start-Job -ScriptBlock{
    
    $Computer = $args[0]
    $RemoteComputerFolder = $args[1]
    $OutPutFile = $args[2]
    
        
    $l = $RemoteComputerFolder.Length - 3
    $RemoteComputerFolder = $RemoteComputerFolder.Substring(3,$l) #Remove c:\ from the string
    
    $RemoteWorkSpace = "\\"+$Computer+"\c$\"+$RemoteComputerFolder
    $RemoteLink = "\\"+$Computer+"\c$\ProgramData\Start Menu\Programs\Startup\getPrn.lnk"
    $RemoteDataFile = "\\"+$Computer+"\c$\"+$RemoteComputerFolder+"\DefaultPrinter.txt"
    
    if(Test-Path -Path $RemoteDataFile){
      $Data = Get-Content -Path $RemoteDataFile
      $OutString = $Computer+","+$Data
      Out-File -Append -FilePath $OutPutFile -InputObject $OutString
      
      if (Test-Path -Path ($RemoteComputerFolder)){Remove-Item -Path $RemoteComputerFolder -Recurse -Force}
      
      #Remote CleanUp
      Remove-Item -Path $RemoteWorkSpace -Recurse -Force
      Remove-Item -Path $RemoteLink -Recurse -Force
      
      return $true
    }else{
      $OutString = $Computer+",Push Succeded"
      Out-File -Append -FilePath $OutPutFile -InputObject $OutString
      return $false
    }

  } -ArgumentList $Computer, $RemoteComputerFolder, $OutPutFile

}

Clear-Host
# User Defined Variables:
[string]$ComputerList = 'C:\Work\ComputerList.csv'
[timespan]$CyclePause = New-TimeSpan -Minutes 15 -Seconds 00
[string]$RemoteComputerFolder = "c:\2r23edgtrg67u8iyhjrtegw231refqw2356yh"
[int]$processNumberMax = 5
[string]$DataCompilationFile = 'C:\Work\DataCompilation.csv'

#Internal Variables
[bool]$TaskIsNotCompleted = $true
[string]$OutPutFile = 'C:\Work\WorkInProgress.csv'

do{

    if (Test-Path -Path $ComputerList) { [Object[]]$Computers = Get-Content -Path $ComputerList} else {Write-Output -InputObject "Computer list file not present or readable:", $ComputerList;break}
      
    $TaskIsNotCompleted = $false
    Write-Output -InputObject "Starting Cycle"
    foreach ($Computer in $Computers){

        $LogData = $Computer

        #Check if ',' do not exists on the line
        if(-not ($Computer -like '*,*')){
            
            $LogData = $LogData+" - Dont have , - Starting Push"
            $TaskIsNotCompleted = $true
            Push-Scrip -Computer $Computer -RemoteComputerFolder $RemoteComputerFolder -processNumberMax $processNumberMax -OutPutFile $OutPutFile
            
        }elseif ($Computer -like '*,Push Succeded*') {
          
          $LogData = $LogData+" - Script pushed in previous cycle, starting Get"
          $Computer = $Computer -replace ",.*" #remove everything after ,
          
          #Starting Get
          #Get-Data is truen when Data File was found
          if( Get-Data -Computer $Computer -RemoteComputerFolder $RemoteComputerFolder -processNumberMax $processNumberMax -OutPutFile $OutPutFile){
            $TaskIsNotCompleted = $false

          }else{

            $TaskIsNotCompleted = $true
          }
          
        }else{
          Out-File -Append -FilePath $DataCompilationFile -InputObject $Computer
        }
        
        Write-Output -InputObject $LogData

    }

    #Whaits until there is no more bacground jobs
    do {
    
        Start-Sleep -Seconds 1

        #Remove Completed Jobs
        foreach ($job in Get-Job) {
            if ( $job.State -eq "Completed" ) { Remove-Job -Id $job.Id }
        }
        
        #Update number of active processes
        $activeProcesses = (Get-Job).Count

    }until ($activeProcesses -eq 0)


    if (Test-Path -Path $ComputerList) { Remove-Item -Path $ComputerList}
    if (Test-Path -Path $OutPutFile) { 
      Rename-Item -Path $OutPutFile -NewName $ComputerList
      $TaskIsNotCompleted = $true
    }
    Write-Output -InputObject "Cycle Finished, starting Sleep"
    Write-Output -InputObject "__________________________________________"
    
    # Total time to sleep
    $start_sleep = $CyclePause.TotalSeconds

    # Time to sleep between each notification
    $sleep_iteration = 2
    
    Write-Output ( "Sleeping {0} seconds ... " -f ($start_sleep) )
    for ($i=1 ; $i -le ([int]$start_sleep/$sleep_iteration) ; $i++) {
        Start-Sleep -Seconds $sleep_iteration
        $percent = $sleep_iteration*$i*100/$start_sleep
        Write-Progress -PercentComplete $percent -CurrentOperation ("Sleep {0}s" -f ($start_sleep)) ( " {0}s ..." -f ($i*$sleep_iteration)  )
    }
    Write-Progress -CurrentOperation ("Sleep {0}s" -f ($start_sleep)) -Completed "Done waiting for X to finish"  
  
    #Start-Sleep -Seconds $CyclePause.TotalSeconds

}while ($TaskIsNotCompleted)

