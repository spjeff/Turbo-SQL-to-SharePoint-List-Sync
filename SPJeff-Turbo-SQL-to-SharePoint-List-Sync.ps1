<#
.SYNOPSIS
    Read source SQL table and synchronize into SharePoint Online (SPO) List destination.
.DESCRIPTION
    Turbo sync SQL table to SharePoint Online List.  Optimized for speed. Insert, Update, and Delete to match rows and columns.  Great for PowerPlatform, PowerApps, and PowerBI integration scenarios.  Simple PS1 to run from on-premise virtual machine Task Scheduler.  No need for Data Managment Gateway (DMG) or complex firewall rules.  Optimized with usage of primary key, index columns, [Compare-Object] cmdlet [System.Data.DataRow] type, hashtable, and PNP batch HTTP POST network traffic.  Includes transcript LOG for support.

    * Script will delete rows from the SharePoint List that are not in the SQL table.  
    * Script will add rows to the SharePoint List that are not in the SharePoint List.  
    * Script will update rows in the SharePoint List that are different in the SQL table.
    * Script will not update rows in the SharePoint List that are the same in the SQL table.
.EXAMPLE
    C:\> .\SPJeff-Turbo-SQL-to-SharePoint-List-Sync.ps1
.NOTES
    File Name:  SPJeff-Turbo-SQL-to-SharePoint-List-Sync.ps1
    Version:    1.0
    Author:     Jeff Jones - @spjeff
    Modified:   2021-02-12
.LINK
    https://github.com/spjeff/Turbo-SQL-to-SharePoint-List-Sync
#>

# DotNet Assembly
using namespace System.Data
using namespace System.Collections.Generic

function Write-Yellow($message) {
    Write-Host $message -ForegroundColor "Yellow"
}

# PowerShell Modules
# from https://stackoverflow.com/questions/28740320/how-do-i-check-if-a-powershell-module-is-installed
Write-Yellow "Loading PowerShell modules..."
@("SQLServer", "PNP.PowerShell") | ForEach-Object {
    if (!(Get-Module -ListAvailable -Name $_)) {
        Write-Yellow "Installing module: $_"
        Install-Module -Name $_ -Force
    }
    else {
        Write-Yellow "Loading module: $_"
        Import-Module $_ -ErrorAction "SilentlyContinue" -WarningAction "SilentlyContinue" | Out-Null
    }
}

# Helper functions
# from https://www.powershellgallery.com/packages/ConvertFrom-DataRow/0.9.1/Content/ConvertFrom-DataRow.psm1
function ConvertFrom-DataRow {
    [CmdletBinding( DefaultParameterSetName = 'AsHashtable' )]
    [OutputType( [System.Collections.Specialized.OrderedDictionary], ParameterSetName = 'AsHashtable' )]
    [OutputType( [pscustomobject], ParameterSetName = 'AsObject' )]
    param(
        [Parameter( Mandatory, Position = 0, ValueFromPipeline = $true )]
        [DataRow[]]
        $InputObject,
        [Parameter( Mandatory, ParameterSetName = 'AsObject' )]
        [switch]
        $AsObject,
        $DbNullValue = $null
    )
    begin {
        [List[string]]$Columns = @()
    }
    process {
        foreach ( $DataRow in $InputObject ) {
            if ( $Columns.Count -eq 0 ) {
                $Columns.AddRange( [string[]]$DataRow.Table.Columns.ColumnName )
            }
            $ReturnObject = @{}
            $Columns | ForEach-Object {
                if ( $DataRow.$_ -is [System.DBNull] ) {
                    $ReturnObject[$_] = $DbNullValue
                }
                else {
                    $ReturnObject[$_] = $DataRow.$_
                }
            }
            if ( $AsObject ) {
                Write-Output ( [pscustomobject]$ReturnObject )
            }
            else {
                Write-Output ( $ReturnObject )
            }
        }
    }
}

function ProcessSQLtoSPLISTSync($xmlSource, $xmlDestination, $xmlMapping) {
    
    # STEP 1 - Connect SQL
    $sqlQuery = $xmlMapping.query
    $sqlServer = $xmlDestination.server
    $sqlDatabase = $xmlDestination.database
    $sqlUser = $xmlDestination.username
    $sqlPass = $xmlSource.password
    $sqlPrimaryKey = $xmlMapping.primarykey
    $sqlSource = Invoke-Sqlcmd -Query $sqlQuery -ServerInstance $sqlServer -Database $sqlDatabase -Username $sqluser -Password $sqlpass
    $sqlSourceHash = $sqlSource | ConvertFrom-DataRow

    # STEP 2 - Connect SPO
    $spUrl = $xmlDestination.url
    $spClientId = $xmlDestination.clientid
    $spClientSecret = $xmlDestination.clientsecret
    $spListName = $xmlDestination.list
    $spMatchItem = $null

    # Dynamic schema.  SPLIST always has [Id] and [Title] fields.  Append SQL columns to SPLIST fields
    $spFields = "Id", $sqlPrimaryKey
    $sqlSource[0].Table.Columns | Where-Object { $_.ColumnName -ne $sqlPrimaryKey } | ForEach-Object { $spFields += $_.ColumnName }

    # Connect to SPO and get SPLIST items with dynamic schema
    Connect-PnPOnline -Url $spUrl -ClientId $spClientId -ClientSecret $spClientSecret -WarningAction "Silentlycontinue"
    $spDestination = Get-PnPListItem -List $spListName -Fields $spFields -PageSize "4000"

    # Measure changes to SPLIST
    $added = 0
    $updated = 0
    $deleted = 0
    
    # Measure SPLIST rows before and after
    $beforeCount = $spDestination.Count
    $afterCount = $sqlSource.Count

    # STEP 3 - Delete excess SPLIST items on destination by comparing primary keys
    $spBatch = New-PnPBatch
    foreach ($item in $spDestination) { 
        # Primary Key
        $pk = $item[$sqlPrimaryKey]

        # Loop comparison
        if ($sqlSource.$sqlPrimaryKey -notcontains $pk) {
            # Delete row
            Remove-PnPListItem -List $spListName -Identity $item.Id -Batch $spBatch
            $deleted++
            Write-Yellow "Deleted: $pk"
        }
    }
    Invoke-PnPBatch -Batch $spBatch -Force

    # STEP 4 - Add or update SPLIST items on destination by comparing primary keys
    $spBatch = New-PnPBatch
    foreach ($row in $sqlSourceHash) {
        # Flatten SQL row to hashtable
        $hash = ($row | Select-Object ($spFields | Select-Object -Skip 1))

        # Primary key search, if row exists
        $pk = $row[$sqlPrimaryKey]
        $spMatchItem = $spDestination | Where-Object { $_[$sqlPrimaryKey] -eq $pk }

        # First matched SPLIST row only
        if ($spMatchItem -is [System.Array]) {
            $spMatchItem = $spMatchItem[0]
        }

        # Format objects consistently for [Compare-Object] support
        if ($spMatchItem) {
            $hashsp = ([Hashtable]$spMatchItem.FieldValues) | Select-Object ($spFields | Select-Object -Skip 1)
            $hashsp."$sqlPrimaryKey" = [int]$hashsp."$sqlPrimaryKey"
        }

        # If row does not exist, add it
        if (!$spMatchItem) {
            # Insert row
            $item = Add-PnPListItem -List $spListName -Values $row -Batch $spBatch
            $added++
            Write-Yellow "Added: $sqlPrimaryKey = $pk"
        }
        else {
            # If row does exist, update it
            $needUpdate = (ConvertTo-Json $hash) -ne (ConvertTo-Json $hashsp)
            if ($needUpdate) {
                # Update row
                Write-Yellow "Updated: $pk"
                Set-PnPListItem -List $spListName -Identity $spMatchItem.Id -Values $row -Batch $spBatch
                $updated++
                $needUpdate = $false
            }
        }
    }

    # Invoke HTTP POST network request
    Invoke-PnPBatch $spBatch

    # Display results with SQL and SPLIST row counts
    Write-Yellow "Source Rows        = $($beforeCount)"
    Write-Yellow "Destination Rows = $($afterCount)"
    Write-Yellow "---"
    Write-Yellow "Added      = $added"
    Write-Yellow "Updated    = $updated"
    Write-Yellow "Deleted    = $deleted"
}

# Main application
function Main() {
    # Load XML config and loop through each mapping
    [xml]$config = Get-Content "SPJeff-Turbo-SQL-to-SharePoint-List-Sync.xml"
    foreach ($mapping in $config.config.mappings) {
        # Get source and destination from config
        $source = $config.config.sources | Where-Object { $_.name -eq $mapping.source }
        $destination = $config.config.destinations | Where-Object { $_.name -eq $mapping.destination }

        # Process SQL source to SPLIST destination sync
        ProcessSQLtoSPLISTSync $source $destination $mapping
    }
}

# Open LOG with script name
$prefix = $MyInvocation.MyCommand.Name
$host.UI.RawUI.WindowTitle = $prefix
$stamp = Get-Date -UFormat "%Y-%m-%d-%H-%M-%S"
Start-Transcript "$PSScriptRoot\log\$prefix-$stamp.log"
$start = Get-Date
Main

# Close LOG and display time elapsed
$end = Get-Date
$totaltime = $end - $start
Write-Yellow "`nTime Elapsed: $($totaltime.tostring("hh\:mm\:ss"))"
Stop-Transcript