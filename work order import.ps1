# This is a script to import WorkOrder data from the IFS share on the iSeries to the WorkOrder table in SharePoint for the Production Workflow System.
#  Change log: 
#   v0.1. ESnow. 11/05/2016. Initial build.
#   v0.2. ESnow. 12/05/2016. Added logging and checks for Z: drive mapping
#   v0.3. ESnow. 12/05/2016. Added optional SQL credentials or Windows pass-thru by checking for the $dbUser variable
#   v0.4. ESnow. 13/05/2016. Fixed driving mapping error
#   v0.5. ESnow. 13/05/2016. Scrapped drive mapping method and implemented FTP by a way of connecting to the IFS
#   v0.6. ESnow. 13/05/2016. Added as scheduled task on MDM-SERV
#   v0.7. ESnow. 14/05/2016. Added switchable debug to log file feature
#   v0.8. ESnow. 14/05/2016. Fixed bug when the first WO .csv wasn't getting processed correctly (added cd c:\ after each invoke-sqlcmd commandlet)
#   v0.9. ESnow. 23/05/2016. Enabled pre-connection to database to fix bug where files were processed but not added to the database.
#   v1.0. ESnow. 27/05/2016. Changed database to the live DB on SPS-SERV

#  NOTES: ESnow. 23/05/2016. I've found that if the processed files in the temp directory are added to the database, they aren't moved to the archive folder immediatly after.
#                            Instead, they are moved at the next time the script is processed. This can make deciphering the logs a little confusing but with the checks that 
#                            are in place, the file isn't added to the DB table again and is moved/deleted as expected at the next scheduled task window.


# ================================================ SET SCRIPT VARIABLES ======================================================================

$debug = 1                    # 1 = on 0 = off
$dbServer = "sps-serv"
$db = "ProdManagementSystem"
$dbPass = ""                  # Only required if using SQL Logins
$dbUser = ""                  # Only required if using SQL Logins
$dbTable = "dbo.WorkOrders"
$date = get-date -format "dd/MM/yyyy HH:mm:ss"
$logFile = "\\fs-3\production workflow system\logs\" + (get-date -format "yyyy-MM-dd") + ".log"
$workOrderArchive = "\\fs-3\production workflow System\archiveFiles\"
cd C:\

# Empty $error variable ##############################################################################################################

$error.Clear()

# ================================================ USE FTP TO DOWNLOAD .CSV FROM IFS =========================================================

# This points to an .ftp file which has the commands necessary to download any .csv file from the directory. #########################

if (Test-Connection -ComputerName manchester -Count 1) {

    start-job -scriptblock {ftp -i -s:"\\fs-2\infotech\eddy\scripts\Production Workflow System\connect.ftp"}

    }

# Wait for ftp process to complete before continuing

while (get-job -State "running") {

    start-sleep -Seconds 2

}

if ($debug -eq 1) {

    add-content -Path $logFile -Value "$date - DEBUG   - Downloaded files to the temp drive"

}

$sourceWorkOrdersPath = "\\fs-3\production workflow System\temp\*.csv"
$sourceWorkOrders = Get-ChildItem -path $sourceWorkOrdersPath

# ================================================ CHECK IF LOG FILE EXISTS ==================================================================

if (!(Test-Path $logFile)) {
    
    New-Item $logFile

    }

# ================================================ CHECK FOR FILES  ==========================================================================

# If there are no new files to process, then exit out of the script immediatly. #######################################################

if (!(Test-Path $sourceWorkOrdersPath)) {

    $date = get-date -format "dd/MM/yyyy HH:mm:ss"
    add-content -Path $logFile -Value "$date - INFO    - No new Work Orders to process - Exiting script"
    Exit

    }

# Create connection to the database. ##################################################################################################

$dbConnect = Invoke-Sqlcmd -Query "SELECT TOP 2 * FROM dbo.WorkOrders" -server $dbServer -Database $db

if ($debug -eq 1) {

    if ($dbConnect) {

        add-content -Path $logFile -Value "$date - DEBUG   - Connected to the database"

    } else {

        add-content -Path $logFile -Value "$date - DEBUG   - Nothing found in database! Check connection"

    }

} 

cd c:\

# ================================================ DATA INPUT INTO THE DB ====================================================================


foreach ($wo in $sourceWorkOrders) {

# Clear $error variable to provide useful logging results

    $error.clear()

# Serialise the data from the WorkOrder

    $importedWo = Import-Csv $wo

    $WorkOrder = $importedWo.WORKORDER
    $ProductCode = $importedWo.PRODUCTCODE
    $ToolNumber = $importedWo.TOOLNUMBER
    $MachineNumber = $importedWo.MACHINENUMBER
    $Status = $importedWo.STATUS
    $ToolCavities = $importedWo.TOOLCAVITIES
    $WorkOrderFileName = $wo.name

    if ($debug -eq 1) {
    
        $date = get-date -format "dd/MM/yyyy HH:mm:ss"
        add-content -Path $logFile -Value "$date - DEBUG   - Processing $WorkOrderFileName"
        add-content -Path $logFile -Value "$date - DEBUG   - Archive folder: $workOrderArchive"
        add-content -Path $logFile -Value "$date - DEBUG   - Work Order: $wo"
        add-content -Path $logFile -Value "$date - DEBUG   - Work Order file check: $workOrderArchive$WorkOrderFileName"
        Add-Content -Path $logFile -Value "$date - DEBUG   - Work Order Number in .csv: $WorkOrder"

    }
    
# Check for duplicates to prevent duplicated PrimaryKey errors in the SQL INSERT command #############################################################################
# Also check if the $dbuser variable is empty - if it is then query using Windows credentials, else use SQL credentials
    
    if ($dbUser -eq "") {
    
        $dupCheck = Invoke-Sqlcmd -Query "SELECT dbo.WorkOrders.WorkOrder FROM dbo.WorkOrders WHERE dbo.WorkOrders.WorkOrder = '$WorkOrder'" -server $dbServer -Database $db
        cd c:\

        if ($debug -eq 1) {

            $date = get-date -format "dd/MM/yyyy HH:mm:ss"
            add-content -Path $logFile -Value "$date - DEBUG   - No SQL logins specified - using Windows pass-through"

        }

    } else {

        $dupCheck = Invoke-Sqlcmd -Query "SELECT dbo.WorkOrders.WorkOrder FROM dbo.WorkOrders WHERE dbo.WorkOrders.WorkOrder = '$WorkOrder'" -server $dbServer -Database $db -Username $dbUser -Password $dbPass
        cd c:\

        if ($debug -eq 1) {

            $date = get-date -format "dd/MM/yyyy HH:mm:ss"
            add-content -Path $logFile -Value "$date - DEBUG   - No SQL logins specified - using Windows pass-through"

        }

    }

    if ($debug -eq 1) {
    
        $date = get-date -format "dd/MM/yyyy HH:mm:ss"
        add-content -Path $logFile -Value "$date - DEBUG   - dupCheck: $dupCheck"

    }

    if ($dupCheck) {

        if ($debug -eq 1) {
        
            $date = get-date -format "dd/MM/yyyy HH:mm:ss"
            add-content -Path $logFile -Value "$date - DEBUG   - dupCheck: True - $WorkOrderFileName found in database"

        }

        $date = get-date -format "dd/MM/yyyy HH:mm:ss"
        add-content -Path $logFile -Value "$date - WARNING - Processed filename: $WorkOrderFileName - WorkOrder: '$WorkOrder' - Already exists in the database"

# Check if the .csv on the temp directory has already been archived. If it hasn't, then attempt to archive now ##################################################################

        if (!(Test-Path "$workOrderArchive$WorkOrderFileName")) {

            try {

                Move-Item -Path $wo -Destination $workOrderArchive -ErrorAction stop
                $date = get-date -format "dd/MM/yyyy HH:mm:ss"
                add-content -Path $logFile -Value "$date - INFO    - Moved WorkOrder File: '$wo' to archive - $workOrderArchive"

                } catch {

                foreach ($line in $error[0]) {

                    $errorLine = $line[0]

                }

                $date = get-date -format "dd/MM/yyyy HH:mm:ss"
                Add-Content -Path $logFile -Value "$date - ERROR   - UNABLE TO MOVE FILE: '$wo' TO ARCHIVE - $errorLine"

                }

            } else {

# If the .csv on the temp directory is already archived, then attempt to delete it from the IFS #################################################################################

                try {

                    Remove-Item -Path $wo -Force -ErrorAction stop
                    $date = get-date -format "dd/MM/yyyy HH:mm:ss"
                    add-content -Path $logFile -Value "$date - INFO    - Deleted WorkOrder File: '$workOrderFileName' from the temp directory - Already exists in archive: $workOrderArchive"

                    } catch {

                    foreach ($line in $error[0]) {

                        $errorLine = $line[0]

                    }

                    $date = get-date -format "dd/MM/yyyy HH:mm:ss"
                    add-content -Path $logFile -Value "$date - ERROR   - UNABLE TO DELETE FILE: '$workOrderFileName' from the temp directory - $errorLine"

                    }

            }

    } else {

# If the content in the .csv is NOT a duplicate Work Order, then add the data to the dbo.WorkOrders table #############################################################
# Also check if the $dbuser variable is empty - if it is then query using Windows credentials, else use SQL credentials

        if ($debug -eq 1) {
        
            $date = get-date -format "dd/MM/yyyy HH:mm:ss"
            add-content -Path $logFile -Value "$date - DEBUG   - dupCheck: False - $WorkOrderFileName not found in database"

        }

        if ($dbUser -eq "") {

            Invoke-Sqlcmd -Query "
                INSERT INTO dbo.WorkOrders (
                WorkOrder, 
                ProductCode, 
                ToolNumber, 
                MachineNumber, 
                Status, 
                ToolCavities) 
            VALUES (
                '$WorkOrder', 
                '$ProductCode', 
                '$ToolNumber', 
                '$MachineNumber', 
                '$Status', 
                '$ToolCavities')" -server $dbServer -database $db
                cd c:\

        } else {

            Invoke-Sqlcmd -Query "
                INSERT INTO dbo.WorkOrders (
                WorkOrder, 
                ProductCode, 
                ToolNumber, 
                MachineNumber, 
                Status, 
                ToolCavities) 
            VALUES (
                '$WorkOrder', 
                '$ProductCode', 
                '$ToolNumber', 
                '$MachineNumber', 
                '$Status', 
                '$ToolCavities')" -server $dbServer -database $db -username $dbUser -password $dbPass
                cd c:\

        }

        $date = get-date -format "dd/MM/yyyy HH:mm:ss"
        add-content -Path $logFile -Value "$date - ADD     - Processed WorkOrder: '$WorkOrder' - Added to the database"

        try {

            Move-Item -Path $wo -Destination $workOrderArchive -Force
            $date = get-date -format "dd/MM/yyyy HH:mm:ss"
            add-content -Path $logFile -Value "$date - INFO    - Moved WorkOrder File: '$wo' to archive - $workOrderArchive"

            } catch {

            $date = get-date -format "dd/MM/yyyy HH:mm:ss"
            Add-Content -Path $logFile -Value "$date - ERROR   - UNABLE TO MOVE FILE: '$wo' TO ARCHIVE!"

            }

    }

    cd c:\

}

Add-Content -Path $logFile -Value "------------------------------------------------------------------------------"

exit
