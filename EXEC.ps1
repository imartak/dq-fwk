param (
    [string]$runFile = "run.txt",
    [string]$busDate
 )
 
###########################################################################
# Set Location to current folder

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
cd $scriptPath
 
###########################################################################
# Load config.txt variables

Get-Content ".\config.txt" | Foreach-Object -begin {$h=@{}} -process { 
    $k = [regex]::split($_,'=');
    if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $h.Add($k[0], $k[1]) } 
}

###########################################################################
# Define constnants

$sqlplus = $h.sqlplus

$global:C_SOURCE_SCHEMA_DWH = 1
$global:C_SOURCE_SCHEMA_VSDPB = 2
$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($False)

if($busDate){
    $bus_date = $busDate
} else {
    $bus_date = $h.bus_date
}

$cur_date = "undefined"
$err_date = "undefined"

###########################################################################
# Import DB connections

$connections = Import-Csv ("connections.txt")

###########################################################################
# Define functions

function log ([string] $message="") {
    $runName = $runFile.Split(".",2, [System.StringSplitOptions]::RemoveEmptyEntries)[0]
    $exec_log_file = "logs\" + (Get-Date -format "yyyy-MM-dd") + "_" + $runName + "_exec.log"
    $sql_log_file = "logs\" + (Get-Date -format "yyyy-MM-dd") + "_" + $runName + "_sql.log"
    $log_message = (Get-Date -format "yyyy-MM-dd HH:mm:ss") + " " + $message
    Add-Content $exec_log_file $log_message
    Add-Content $sql_log_file $log_message
    Write-Host $log_message
}

function get-connection ([int] $schema_id){
    foreach($con in $connections){
        if($con.schema_id -eq $schema_id){
            return $con
        }
    }
    return $null
}

function Load-Assemblies(){
    log ("Loading assemblies")
    $asm_local_path = $scriptPath + "\.assembly"
    
    if(![System.IO.File]::Exists($asm_local_path)){
        $processor = $ENV:PROCESSOR_ARCHITECTURE.ToLower()
        $asm_folder = (Get-ChildItem 'C:\Windows\winsxs\' | Where-Object {$_.Name -like ($processor+'_system.web*')} | Sort-Object -Property 'LastWriteTime' -Descending | Select-Object -First 1)
        $asm_file = "C:\Windows\winsxs\" + $asm_folder.Name + "\System.Web.dll"
        $asm_output = [System.Reflection.Assembly]::LoadFrom($asm_file)
        $asm_output.Location | Out-File -FilePath $asm_local_path -Encoding "Default"
        $file = Get-Item $asm_local_path -Force
        $file.Attributes = "Archive","Hidden"
    } else {
        $asm_location = Get-Content $asm_local_path
        $asm_output = [System.Reflection.Assembly]::LoadFrom($asm_location)
    }
    log ("Assembly added: " + $asm_output.Location)
    $env:NLS_LANG="CZECH_CZECH REPUBLIC.EE8MSWIN1250"
}

function control-log-start([string] $sql_file,[string] $log_file){
    $control_log_temp_file = ("DQ_CONTROL_LOG_Insert_" + $runFile)
    Get-Content ("templates\DQ_CONTROL_LOG_Insert.sql") | Add-Content $control_log_temp_file
    replace-str "\[runFile\]" $runFile $control_log_temp_file $control_log_temp_file
    replace-str "\[sql_file\]" $sql_file $control_log_temp_file $control_log_temp_file
    replace-str "\[log_file\]" $log_file $control_log_temp_file $control_log_temp_file
    replace-str "\[bus_date\]" $bus_date $control_log_temp_file $control_log_temp_file
    $connect = get-connection $global:C_SOURCE_SCHEMA_DWH
    $args = "/c echo @" + $control_log_temp_file + " | " + $sqlplus + " " + $connect.username + "/" + $connect.password + "@" + $connect.con_string + " > nul "
    cmd $args    
    $control_log_key = (Get-Content ("currval_" + $runFile)).trim()
    Remove-Item $control_log_temp_file
    Remove-Item ("currval_" + $runFile)
    return $control_log_key
}

function control-log-end([string] $control_log_key, [string] $status){
    $control_log_temp_file = ("DQ_CONTROL_LOG_Update_" + $runFile)
    Get-Content ("templates\DQ_CONTROL_LOG_Update.sql") | Add-Content $control_log_temp_file
    replace-str "\[key\]" $control_log_key $control_log_temp_file $control_log_temp_file
    replace-str "\[status\]" $status $control_log_temp_file $control_log_temp_file
    $connect = get-connection $global:C_SOURCE_SCHEMA_DWH
    $args = "/c echo @" + $control_log_temp_file + " | " + $sqlplus + " " + $connect.username + "/" + $connect.password + "@" + $connect.con_string + " > nul "
    cmd $args    
    Remove-Item ($control_log_temp_file)
}

function run-sql ([string] $path, [int] $source_schema_id) {
    log ("Running " + $path)
    $runName = $runFile.Split(".",2, [System.StringSplitOptions]::RemoveEmptyEntries)[0]
    $sql_log_file = "logs\" + (Get-Date -format "yyyy-MM-dd") + "_" + $runName + "_sql.log"
    $control_log_key = control-log-start $path $sql_log_file
    $connect = get-connection $source_schema_id
    $args = "/c echo @" + $path + " | " + $sqlplus + " " + $connect.username + "/" + $connect.password + "@" + $connect.con_string ## + " >> " + $sql_log_file
    $output = cmd $args
    Add-Content $sql_log_file $output
    $str = Out-String -InputObject $output
    if($str.IndexOf("ERROR ") -ne -1){
        control-log-end $control_log_key "NOK"
    } else {
        control-log-end $control_log_key "OK"
    }
}

function replace-str ([string] $old_string, [string] $new_string, [string] $src_path, [string] $dest_path){
    (Get-Content $src_path) | Foreach-Object {
        $_ -replace $old_string, $new_string
    } | Set-Content $dest_path
}

function merge-files ([string] $header, [string] $main, [string] $footer, [string] $output) {
    $new_file = New-Item -ItemType file $output –force
    Get-Content $header | Add-Content $output
    Get-Content $main | Add-Content $output
    Get-Content $footer | Add-Content $output
    log ("File created: " + $new_file)
}

function prepare-dqc([string] $filename){
    (gc (".\DQI_PREPARED\" + $filename)) | ? {$_.trim() -ne "" } | set-content (".\DQI_PREPARED\" + $filename)
}

function set-bus-date([string] $bus_date){
    $date = [datetime]::ParseExact($bus_date,'yyyyMMdd',$null)
    
    $out_load_partition = 'B' + $bus_date
    $cur_date = (Get-Date $date -format 'dd.MM.yy')
    
    Get-Content ("templates\DQ_CURRENT_DATE_Update.sql") | Add-Content ("DQ_CURRENT_DATE_Update.sql")
    replace-str "\[CUR_DATE\]" $cur_date ("DQ_CURRENT_DATE_Update.sql") ("DQ_CURRENT_DATE_Update.sql")
    replace-str "\[OUT_LOAD_PARTITION\]" $out_load_partition ("DQ_CURRENT_DATE_Update.sql") ("DQ_CURRENT_DATE_Update.sql")
    run-sql ("DQ_CURRENT_DATE_Update.sql") $global:C_SOURCE_SCHEMA_DWH
    Remove-Item ("DQ_CURRENT_DATE_Update.sql")
}

function dwh-run ([string] $dq_dqi_key, [string] $filename) {
    $id = $filename.Split(".",2, [System.StringSplitOptions]::RemoveEmptyEntries)[0]

    # Preparing DQ control for merging
    log ("Preparing SQL file DQI_PREPARED\" + $filename)
    replace-str ";" "" (".\DQI\" + $filename) (".\DQI_PREPARED\" + $filename)
    # hack
    prepare-dqc $filename
    
    # Computing RESULTs of DQ control
    log ("Creating runnable SQL file DQI_RESULT\" + $filename)
    merge-files "templates\header_result.txt" ("DQI_PREPARED\" + $filename) "templates\footer_result.txt" ("DQI_RESULT\" + $filename)
    replace-str "\[bus_date\]" $bus_date ("DQI_RESULT\" + $filename) ("DQI_RESULT\" + $filename)
    replace-str "\[dq_dqi_key\]" $dq_dqi_key ("DQI_RESULT\" + $filename) ("DQI_RESULT\" + $filename)     
    run-sql ("DQI_RESULT\" + $filename) $global:C_SOURCE_SCHEMA_DWH
    
    log ("Creating runnable SQL file DQI_ERROR_LIST\" + $filename)
    merge-files "templates\select_header_error_list.txt" ("DQI_PREPARED\" + $filename) "templates\select_footer_error_list.txt" ("DQI_ERROR_LIST\" + $filename)
    replace-str "\[id\]" $id ("DQI_ERROR_LIST\" + $filename) ("DQI_ERROR_LIST\" + $filename)
    replace-str "\[bus_date\]" $bus_date ("DQI_ERROR_LIST\" + $filename) ("DQI_ERROR_LIST\" + $filename)
    replace-str "\[dq_dqi_key\]" $dq_dqi_key ("DQI_ERROR_LIST\" + $filename) ("DQI_ERROR_LIST\" + $filename)
    run-sql ("DQI_ERROR_LIST\" + $filename) $global:C_SOURCE_SCHEMA_DWH
    
    New-PSDrive -Name Y -PSProvider filesystem -Root \\YB0PSA003\2xxx
    $out_directory = 'Y:\2600\20_EBV\ERROR_LIST\DAX'+$err_date
    if(! (Test-Path $out_directory)){
        New-Item $out_directory -type directory 
    }
    $csv_output_name = 'ERROR_LIST_' + $bus_date + '_' + $id + '.csv'
    Copy-Item ('CSV_OUTPUT\'+$csv_output_name) ($out_directory+'\'+$csv_output_name)
}

function nvl ([string] $expr1, [string] $expr2){
    if(($expr1 -eq "") -or ($expr1 -eq $null)){
        return $expr2
    } else {
        return "'" + $expr1 + "'"
    }
}

function db-to-dwh-create-select ([string] $dq_dqi_key, [string] $id, [string] $subject) {
    log ("Creating runnable SQL file DQI_" + $subject + "\" + $id + "_SELECT.sql")
    merge-files ("templates\vsdpb_select_header_" + $subject.ToLower() + ".txt") ("DQI_PREPARED\" + $id + ".sql") ("templates\vsdpb_select_footer_" + $subject.ToLower() + ".txt") ("DQI_" + $subject + "\" + $id + "_SELECT.sql")
    replace-str "\[dq_dqi_key\]" $dq_dqi_key ("DQI_" + $subject + "\" + $id + "_SELECT.sql") ("DQI_" + $subject + "\" + $id + "_SELECT.sql")
    replace-str "\[id\]" $id ("DQI_" + $subject + "\" + $id + "_SELECT.sql") ("DQI_" + $subject + "\" + $id + "_SELECT.sql") 
    replace-str "\[bus_date\]" $bus_date ("DQI_" + $subject + "\" + $id + "_SELECT.sql") ("DQI_" + $subject + "\" + $id + "_SELECT.sql") 
}

function db-to-dwh-create-result-insert ([string] $dq_dqi_key, [string] $id) {
    log ("Creating runnable SQL file DQI_RESULT\" + $id + "_INSERT.sql")
    $results = (Import-Csv ("CSV_OUTPUT\" + $id + "_RESULT.csv") -Header "dq_dqi_key","eff_date","bus_date","tot_count","err_count","succ_rate","tot_volume","err_volume","succ_vol_rate")
    if([System.IO.File]::Exists("DQI_RESULT\" + $id + "_INSERT.sql")){    
        Remove-Item ("DQI_RESULT\" + $id + "_INSERT.sql")
    }
    foreach ($result in $results) {
        Get-Content ("templates\vsdpb_insert_result.txt") | Add-Content ("DQI_RESULT\" + $id + "_INSERT.sql")
        replace-str "\[dq_dqi_key\]" $result.dq_dqi_key ("DQI_RESULT\" + $id + "_INSERT.sql") ("DQI_RESULT\" + $id + "_INSERT.sql")
        replace-str "\[eff_date\]" $result.eff_date ("DQI_RESULT\" + $id + "_INSERT.sql") ("DQI_RESULT\" + $id + "_INSERT.sql")
        replace-str "\[bus_date\]" $result.bus_date ("DQI_RESULT\" + $id + "_INSERT.sql") ("DQI_RESULT\" + $id + "_INSERT.sql")
        replace-str "\[tot_count\]" (nvl $result.tot_count "null") ("DQI_RESULT\" + $id + "_INSERT.sql") ("DQI_RESULT\" + $id + "_INSERT.sql")
        replace-str "\[err_count\]" (nvl $result.err_count "null") ("DQI_RESULT\" + $id + "_INSERT.sql") ("DQI_RESULT\" + $id + "_INSERT.sql")
        replace-str "\[succ_rate\]" (nvl $result.succ_rate "null") ("DQI_RESULT\" + $id + "_INSERT.sql") ("DQI_RESULT\" + $id + "_INSERT.sql")
        replace-str "\[tot_volume\]" (nvl $result.tot_volume "null") ("DQI_RESULT\" + $id + "_INSERT.sql") ("DQI_RESULT\" + $id + "_INSERT.sql")
        replace-str "\[err_volume\]" (nvl $result.err_volume "null") ("DQI_RESULT\" + $id + "_INSERT.sql") ("DQI_RESULT\" + $id + "_INSERT.sql")
        replace-str "\[succ_vol_rate\]" (nvl $result.succ_vol_rate "null") ("DQI_RESULT\" + $id + "_INSERT.sql") ("DQI_RESULT\" + $id + "_INSERT.sql")
    }
}

function db-to-dwh-create-error-list-insert ([string] $dq_dqi_key, [string] $id) {
    log ("Creating runnable SQL file DQI_ERROR_LIST\" + $id + "_INSERT.sql")
    $error_lists = (Import-Csv ("CSV_OUTPUT\" + $id + "_ERROR_LIST.csv") -Header "dq_dqi_key", "bus_date", "eff_date", "ref_id1", "ref_id2", "ref_descr", "ref_descr2", "col_1", "col_2", "col_3", "col_4", "col_5")
    if([System.IO.File]::Exists("DQI_ERROR_LIST\" + $id + "_INSERT.sql")){    
        Remove-Item ("DQI_ERROR_LIST\" + $id + "_INSERT.sql")
    }
    
    foreach ($error_list in $error_lists) {
        $sqlFile = "insert into RUSB_OWNER.DQ_ERROR_LIST (DQ_ERROR_LIST_KEY, DQ_DQI_KEY, BUS_DATE, EFF_DATE, REF_ID1, REF_ID2, REF_DESCR1, REF_DESCR2, COL_1, COL_2, COL_3, COL_4, COL_5)" + "`r`n"
        $sqlFile += "values (RUSB_OWNER.DQ_DQI_RESULT_SEQ.nextval, "
        $sqlFile += $error_list.dq_dqi_key+", "
        $sqlFile += $error_list.bus_date+", "
        $sqlFile += $error_list.eff_date+", "
        $sqlFile += (nvl $error_list.ref_id1 "null")+", "
        $sqlFile += (nvl $error_list.ref_id2 "null")+", "
        $sqlFile += (nvl $error_list.ref_descr "null")+", "
        $sqlFile += (nvl $error_list.ref_descr2 "null")+", "
        $sqlFile += (nvl $error_list.col_1 "null")+", "
        $sqlFile += (nvl $error_list.col_2 "null")+", "
        $sqlFile += (nvl $error_list.col_3 "null")+", "
        $sqlFile += (nvl $error_list.col_4 "null")+", "
        $sqlFile += (nvl $error_list.col_5 "null")+"); "
        Add-Content ("DQI_ERROR_LIST\" + $id + "_INSERT.sql") $sqlFile
    }
    
    Add-Content ("DQI_ERROR_LIST\" + $id + "_INSERT.sql") "commit;`r`n/"
}

function db-to-dwh-run ([string] $dq_dqi_key, [string] $filename, [int] $schema_id) {
    $id = $filename.Split(".",2, [System.StringSplitOptions]::RemoveEmptyEntries)[0]

    ## Preparing DQ control for merging
    log ("Preparing SQL file DQI_PREPARED\" + $filename)
    replace-str ";" "" (".\DQI\" + $filename) (".\DQI_PREPARED\" + $filename)
    
    ## Computing RESULTs of DQ control
    db-to-dwh-create-select $dq_dqi_key $id "RESULT"
    run-sql ("DQI_RESULT\" + $id + "_SELECT.sql") $schema_id
    
    ## Storing RESULTs of DQ control
    db-to-dwh-create-result-insert $dq_dqi_key $id
    run-sql ("DQI_RESULT\" + $id + "_INSERT.sql") $global:C_SOURCE_SCHEMA_DWH    
    
    ## Computing ERROR_LISTs of DQ control
    db-to-dwh-create-select $dq_dqi_key $id "ERROR_LIST"
    run-sql ("DQI_ERROR_LIST\" + $id + "_SELECT.sql") $schema_id

}

function split-file ([string] $file){
    log ("Splitting " + $file)    
    $full_path = [System.IO.Path]::GetFullPath($file)

    if([System.IO.File]::Exists($full_path)){
        $count = Get-Content $full_path | Measure-Object –Line
        $size = [Math]::Ceiling($count.Lines / 3)
                
        $file_part = $file.Split(".",2, [System.StringSplitOptions]::RemoveEmptyEntries)
        
        if($count.Lines -eq 0){
            return 0
        } elseif ($count.Lines -lt 6) {
            (Get-Content $full_path) | Out-File ($file_part[0] + "1." + $file_part[1])
            return 1
        } elseif ($count.Lines -lt 11) {
            (Get-Content $full_path)[0..($size-1)] | Out-File ($file_part[0] + "1." + $file_part[1])
            (Get-Content $full_path)[$size..$count.Lines] | Out-File ($file_part[0] + "2." + $file_part[1])
            return 2
        } elseif ($count.Lines -lt 60) {
            (Get-Content $full_path)[0..($size-1)] | Out-File ($file_part[0] + "1." + $file_part[1])
            (Get-Content $full_path)[$size..(2*$size-1)] | Out-File ($file_part[0] + "2." + $file_part[1])
            (Get-Content $full_path)[(2*$size)..$count.Lines] | Out-File ($file_part[0] + "3." + $file_part[1])
            return 3
        } else {
            $size_t = [Math]::Ceiling($count.Lines / 12)
            (Get-Content $full_path)[0..(3*$size_t-1)] | Out-File ($file_part[0] + "1." + $file_part[1])
            (Get-Content $full_path)[(3*$size_t)..(6*$size_t-1)] | Out-File ($file_part[0] + "2." + $file_part[1])
            (Get-Content $full_path)[(6*$size_t)..(9*$size_t-1)] | Out-File ($file_part[0] + "3." + $file_part[1])
            (Get-Content $full_path)[(9*$size_t)..$count.Lines] | Out-File ($file_part[0] + "4." + $file_part[1])
            return 4
        }
    }
}

function move-to-history ([string] $file){
    $file_name = $file.Split(".",2, [System.StringSplitOptions]::RemoveEmptyEntries)[0]
    $new_name = (Get-Date -format "yyyy-MM-dd-HH-mm-ss") + "_" + $file_name + ".txt"
    $full_path = [System.IO.Path]::GetFullPath($file)
    if([System.IO.File]::Exists($full_path)){
        Rename-Item -path $file -newname $new_name
        Move-Item $new_name "history\"   
        log ("File archived: " + $file)
    }    
    
}

function start-run ([string] $run_file) {
    $full_path = [System.IO.Path]::GetFullPath($run_file)

    if([System.IO.File]::Exists($full_path)){
        $reader = [System.IO.File]::OpenText($full_path)
        
        $date = [datetime]::ParseExact($bus_date,'yyyyMMdd',$null)
        $cur_date = (Get-Date $date -format 'dd.MM.yy')
        $err_date = (Get-Date $date -format 'yyyyMM')
        
        # $current_date = Import-Csv -Path "current_date_output.csv" -Delimiter "," -Header "CUR_DATE","OUT_LOAD_PARTITION","STRING_DATE" 
        # $cur_date = $current_date.CUR_DATE 
        # $err_date = $current_date.STRING_DATE
        
        try {
            for(;;) {
                $line = $reader.ReadLine()
                if ($line -eq $null) { break }
                # process the line
                $option = [System.StringSplitOptions]::RemoveEmptyEntries
                $dqc = $line.Split(":",3, $option)
                $source_schema_id = $dqc[2].trim()
                
                if($source_schema_id -eq $global:C_SOURCE_SCHEMA_DWH){
                    dwh-run $dqc[0] $dqc[1]
                }
                
                if($source_schema_id -eq $global:C_SOURCE_SCHEMA_VSDPB){
                    db-to-dwh-run $dqc[0] $dqc[1] $source_schema_id
                }
                
            }
        }
        finally {
            $reader.Close()
        }
    }
}

function run (){
    log "Start"

    if($runFile -eq "run.txt"){
        if(!(Test-Path -Path $runFile)){
            run-sql createRun.sql $global:C_SOURCE_SCHEMA_DWH
        }
        
        $jobCount = split-file($scriptPath + "\run.txt")
        $procs = @()
        
        for($i = 1;$i -lt ($jobCount + 1); $i = $i + 1){
            $procs += $(Start-Process powershell -argument (".\exec.ps1 run"+$i+".txt " + $bus_date) -PassThru)
            log ("Run " + $i + " started")
        }
        
        $procs | Wait-Process
    } else {        
        start-run $runFile
    }    
    
    move-to-history $runFile

    log ("Finish")
}

###########################################################################
# PROCESS

Load-Assemblies

run