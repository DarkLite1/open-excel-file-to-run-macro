<#
    .SYNOPSIS
        Open an Excel file and run a macro.

    .DESCRIPTION
        This script can be triggered by a scheduled task.
        Excel needs to be installed for this to work.

    .PARAMETER ExcelFilePath
        The path to the Excel file to open.

    .PARAMETER LogFolderPath
        The folder where the log file will be saved.

    .PARAMETER MacroToExecute
        The name of the macro to execute. If no name is provided, the script
        will wait for Excel to finish. This can be convenient in case of macro's
        that are started automatically when opening Excel.

    .PARAMETER SaveChangesInExcel
        Save changes in Excel file on TRUE, else no changes are saved.

    .PARAMETER MaxWaitTimeSeconds
        Total time to wait for Excel to finish.
#>

Param (
    [Parameter(Mandatory)]
    [String]$ExcelFilePath,
    [String]$LogFolderPath = $PSScriptRoot,
    [String]$MacroToExecute = $null,
    [Boolean]$SaveChangesInExcel = $true,
    [Boolean]$ShowExcelFileAndDisplayAlerts = $false,
    [Int]$MaxWaitTimeSeconds = 30 * 60
)

try {
    $params = @{
        Path      = Join-Path $LogFolderPath ('Log - ' + (Get-Date).ToString('yyyymmdd - HHmmss') + '.txt')
        NoClobber = $true
    }
    Start-Transcript @params

    $Error.Clear()
    $VerbosePreference = 'Continue'
    $ErrorActionPreference = 'Stop'

    #region Create required Excel folder
    $requiredFolder = 'C:\Windows\SysWOW64\config\systemprofile\Desktop'

    if (-not (Test-Path -LiteralPath $requiredFolder -PathType Container)) {
        Write-Verbose "Create required Excel folder '$requiredFolder'"
        $null = New-Item -Path $requiredFolder -ItemType Directory
    }
    #endregion

    Write-Verbose 'Start script'

    #region Verbose parameters
    Write-Verbose "Params:"
    Write-Verbose "- ExcelFilePath '$ExcelFilePath'"
    Write-Verbose "- LogFolderPath '$LogFolderPath'"
    Write-Verbose "- MacroToExecute '$MacroToExecute'"
    Write-Verbose "- SaveChangesInExcel '$SaveChangesInExcel'"
    Write-Verbose "- MaxWaitTimeSeconds '$MaxWaitTimeSeconds'"
    #endregion

    #region Test file exists
    Write-Verbose 'Test Excel file exists'

    if (-not (Test-Path -LiteralPath $ExcelFilePath -PathType Leaf)) {
        throw "File '$ExcelFilePath' not found"
    }
    #endregion

    #region Start Excel
    Write-Verbose 'Open Excel file'

    try {
        $excel = New-Object -ComObject Excel.Application -Verbose:$false
    }
    catch {
        throw "Failed to start Excel app: Excel not installed: $_"
    }
    #endregion

    #region Show Excel app
    $excel.Visible = $ShowExcelFileAndDisplayAlerts
    $excel.DisplayAlerts = $ShowExcelFileAndDisplayAlerts
    #endregion

    #region Open Excel file
    try {
        Write-Verbose 'Open Excel file'
        $workbook = $excel.Workbooks.Open($ExcelFilePath)
    }
    catch {
        throw "Excel to open Excel File '$ExcelFilePath': $_"
    }
    #endregion

    $app = $excel.Application

    if ($MacroToExecute) {
        #region Execute macro
        Write-Verbose 'Execute macro'

        $app.Run($MacroToExecute)
        #endregion
    }
    else {
        #region Wait for any running macro at startup to finish
        $sleepParams = @{
            Seconds = 10
        }

        $totalRunTimeSeconds = 0
        $finished = $false

        do {
            Write-Verbose "Wait $($sleepParams.Seconds) seconds"

            $totalRunTimeSeconds += $sleepParams.Seconds

            Start-Sleep @sleepParams

            if ($app.Ready) {
                $finished = $true
            }
        } while (
            (-not $finished) -and
            ($totalRunTimeSeconds -lt $maxWaitTimeSeconds)
        )

        if ($finished) {
            Write-Verbose 'Excel finished, stop waiting'
        }
        else {
            Write-Warning "Excel not finished in '$maxWaitTimeSeconds' seconds, stopped waiting and aborted process"
        }
        #endregion
    }
}
catch {
    $errorMessage = "Failed: $_"
    Write-Warning $errorMessage
    throw $errorMessage
}
finally {
    #region Close Excel workbook
    try {
        if ($error) {
            Write-Warning 'Errors found, changes will not be saved'
            $SaveChangesInExcel = $false
        }
        if ($SaveChangesInExcel) {
            Write-Verbose 'Save changes'
        }

        Write-Verbose 'Close Excel file'
        $workbook.Close($SaveChangesInExcel)
    }
    catch {
        Write-Warning "Failed to close Excel workbook: $_"
    }
    #endregion

    #region Close Excel file
    try {
        Write-Verbose 'Close Excel'
        Start-Sleep -Seconds 3
        $excel.Quit()
    }
    catch {
        Write-Warning "Failed to close Excel file: $_"
        Write-Verbose 'Kill Excel process forcefully'
        Get-Process -Name 'excel' -EA Ignore | Stop-Process -Force
    }
    #endregion

    #region Clean up
    Write-Verbose 'Clean up Excel objects'
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    #endregion

    Write-Verbose "Script finished"

    Stop-Transcript
}