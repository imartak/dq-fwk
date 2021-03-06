param (
    [string]$periodicity_key = $null
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
# Load required assembly & sql client path

$sqlplus = '"' + $h.sqlplus + '"'

###########################################################################
# Define constnants

$sqlCreateRunFilePath = ".\createRun.sql"
$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($False)

$C_CDT_WORKPACKAGE_ID = 1
$C_CDT_CLUSTER_ID = 2

$dqi_root_page = $h.dqi_root_page
$dqc_root_page = $h.dqc_root_page

###########################################################################
# Define Web client

$web = New-Object Net.WebClient

$rest_api_url = $h.rest_api_url

$spaceName = $h.spaceName
$username = $h.username
$password = $h.password

$web.Headers.Add('Content-Type', 'application/json')
$web.Headers.Add('Media-Type', 'application/json')
$web.Headers.Add('Accept', 'application/json')
$auth = 'Basic ' + [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($username+":"+$password ))
$web.Headers.Add('Authorization', $auth)
$web.Encoding = [System.Text.Encoding]::UTF8

###########################################################################
# Import DB connections

$connections = Import-Csv ("connections.txt")

############################################################################
## Define functions

function log ([string] $message="") {
    $runName = "init"
    $sql_log_file = "logs\" + (Get-Date -format "yyyy-MM-dd") + "_" + $runName + "_sql.log"
    $log_message = (Get-Date -format "yyyy-MM-dd HH:mm:ss") + " " + $message
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

function Check-Connections(){
    $success = $true
    foreach($con in $connections){
        log ("Checking connection " + $con.schema_id)
        $args = "/c echo exit | " + $sqlplus + " -L " + $con.username + "/" + $con.password + "@" + $con.con_string
        [string]$response = cmd $args
        
        if($response.IndexOf("ERROR") -eq -1){
            log ("Connection " + $con.schema_id + " is working")
        } else {
            log ("Connection " + $con.schema_id + " failed: " + $response.SubString($response.IndexOf("ERROR") + 7))
            $success = $false
        }       
    }
    if(!$success){
        $title = "DB Connection error detected"
        $message = "Do you wish to proceed?"
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
            "Only if you are aware of the consequences."
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 
        switch ($result){
            0 { return $true }
            1 { return $false}
        }
    }
    return $success
}

function Load-Assemblies(){
    log ("Loading assemblies")
    $asm_local_path = ".assembly"
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
}

function Remove-Diacritics([string] $s) {
    $s = $s.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = ""
    for ($i = 0; $i -lt $s.Length; $i++) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($s[$i]) -ne    [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            $sb += $s[$i] 
        }
    }
    return $sb
}

function run-sql ([string] $path, [int] $source_schema_id) {
    log ("Running " + $path)
    $runName = "init"
    $sql_log_file = "logs\" + (Get-Date -format "yyyy-MM-dd") + "_" + $runName + "_sql.log"
    $connect = get-connection $source_schema_id
    $args = "/c echo @" + $path + " | " + $sqlplus + " " + $connect.username + "/" + $connect.password + "@" + $connect.con_string + " >> " + $sql_log_file
    cmd $args
}

function ConvertTo-Json20([object] $item){
    add-type -assembly system.web.extensions
    $ps_js=new-object system.web.script.serialization.javascriptSerializer
    return $ps_js.Serialize($item) 
}

function ConvertFrom-Json20([object] $item){ 
    add-type -assembly system.web.extensions 
    $ps_js=new-object system.web.script.serialization.javascriptSerializer
    return $ps_js.DeserializeObject($item) 
}

function Get-Url([string] $page_id, [object] $params = $null){
    $output_url = $rest_api_url + $page_id
    if($params){
        $i = 0
        $params.GetEnumerator() | Foreach-Object {
            if($i -eq 0) {
                $output_url += "?"
            } else {
                $output_url += "&"
            }
            $output_url += $_.Key + "=" + ([System.Web.HttpUtility]::UrlEncode(([System.Web.HttpUtility]::HtmlDecode($_.Value))))
            $i = $i + 1
        }
    }
    return $output_url
}

function Http-Get([string] $url){
    return $web.DownloadString($url)
}

function Http-Put([string] $url, $json){
    $web.Headers.Add('Content-Type', 'application/json')
    return $web.UploadString($url, "PUT", $json )
}

function Get-Page([string] $page_id, [object] $params){
    $response = Http-Get ((Get-Url $page_id $params))
    return (ConvertFrom-Json20 $response)
}

function Put-Page([object] $input_page){
    $page = $input_page
    $ancestors = @()
    $ancestors += @{
        id=$page.ancestors[-1].id;
        type="page"
    }
    $page.ancestors = $ancestors
    $page.version.number = ($page.version.number + 1)
    $page.body.storage.value = [string]$page.body.storage.value
    $json = (ConvertTo-Json20 $page)

    Http-Put (Get-Url $page.id) $json
}

function Get-Leaves([string] $page_id, [object] $params){
    return @(Get-Descendants $page_id $params $true)
}

function Get-Descendants([string] $page_id, [object] $params, [boolean] $onlyLeaves = $false){
    $json_res = (Get-Page $page_id $params)
    $children = @()
    
    if($json_res.children.page.results){   
        $child_results = @()
        $start = 0
        $limit = 10
        for(;;){
            $ch_params = @{
                status = 'current';
        		spaceKey = $spaceName;
        		expand = 'body.storage,children.page,version,ancestors';
                start=$start;
                limit=$limit
        	}
            $json_res_children = (Get-Page ($page_id + "/child/page") $ch_params)
            $child_results += $json_res_children.results
            if($json_res_children.size -eq $limit){
                $start = $start + $limit
            } else {
                break
            }
        }
        if(!$onlyLeaves){
            Write-Host $json_res.id $json_res.title
            $children += $json_res
        }
        
        foreach ($res in $child_results){
            # Write-Host (Get-Url $res.id $ch_params)   
            $children += Get-Descendants $res.id $ch_params $onlyLeaves
        }
    } else {  
        Write-Host $json_res.id $json_res.title      
        return $json_res
    }
    return $children
}

function Get-PageList([string[]] $cnfl_keys){
    $output = @()
    foreach($cnfl_key in $cnfl_keys){
        $page = Get-Page $cnfl_key @{
            status = 'current';
            spaceKey = $spaceName;
            expand = 'body.storage,children.page,version,ancestors'
        }
        $output += $page
        Write-Host $page.id $page.title
    }
    return $output
}

function Get-Metadata-Value([string] $html, [string] $name){
    $ac_param_position = $html.IndexOf('<ac:parameter ac:name="0">' + $name + '</ac:parameter>') 
    if($ac_param_position -eq -1){
        return $ac_param_position
    }
    $start = $html.SubString($ac_param_position).IndexOf('<![CDATA[') + $ac_param_position + 9
    $end_macro = $html.SubString($ac_param_position).IndexOf('</ac:structured-macro>') + $ac_param_position
    if((($start - $ac_param_position - 9) -lt 0) -or (($end_macro + 9) -lt $start)){
        return ""
    }
    $end = $html.SubString($start).IndexOf(']]')
    return $html.SubString($start, $end)
}

function Set-Metadata-Value([string] $html, [string] $name, [string] $value){
    $ac_param_position = $html.IndexOf('<ac:parameter ac:name="0">' + $name + '</ac:parameter>') 
    if($ac_param_position -eq -1){
        log ("ERROR: Metadata value `"" + $name + "`" does not exist.")
        return $html
    }
    $start = $html.SubString($ac_param_position).IndexOf('<![CDATA[') + $ac_param_position + 9
    $end_macro = $html.SubString($ac_param_position).IndexOf('</ac:structured-macro>') + $ac_param_position
    if(($end_macro + 9) -lt $start){
        [string]$output = $html.SubString(0,$end_macro) + '<ac:plain-text-body><![CDATA[' + $value + ']]></ac:plain-text-body>' + $html.SubString($end_macro)
    } else {
        $end = $html.SubString($start).IndexOf(']]') + $start
        [string]$output = $html.SubString(0,$start) + $value + $html.SubString($end)
    }
    return $output
}

function Get-SQL-Value([string] $html){
    $ac_param_position = $html.IndexOf('<ac:parameter ac:name="title">sql</ac:parameter>') 
    if($ac_param_position -eq -1){
        return $ac_param_position
    }
    $start = $html.SubString($ac_param_position).IndexOf('<![CDATA[') + $ac_param_position + 9
    $end_macro = $html.SubString($ac_param_position).IndexOf('</ac:structured-macro>') + $ac_param_position
    if((($start - $ac_param_position - 9) -lt 0) -or (($end_macro + 9) -lt $start)){
        return ""
    }
    $end = $html.SubString($start).IndexOf(']]')
    return $html.SubString($start, $end)
}

function Set-SQL-Value([string] $html, [string] $value){
    $ac_param_position = $html.IndexOf('<ac:parameter ac:name="title">sql</ac:parameter>') 
    if($ac_param_position -eq -1){
        log ("ERROR: Metadata value `"" + $name + "`" does not exist.")
        return $html
    }
    $start = $html.SubString($ac_param_position).IndexOf('<![CDATA[') + $ac_param_position + 9
    $end_macro = $html.SubString($ac_param_position).IndexOf('</ac:structured-macro>') + $ac_param_position
    if(($end_macro + 9) -lt $start){
        [string]$output = $html.SubString(0,$end_macro) + '<ac:plain-text-body><![CDATA[' + $value + ']]></ac:plain-text-body>' + $html.SubString($end_macro)
    } else {
        $end = $html.SubString($start).IndexOf(']]') + $start
        [string]$output = $html.SubString(0,$start) + $value + $html.SubString($end)
    }
    return $output
}

function Get-NextData([string] $html, [string] $tag, [string] $attribute = $null){
    $start_1 = $html.IndexOf("<" + $tag)
    if($start_1 -eq -1){
        return $start_1
    }
    if($attribute){     
        $start = $html.SubString($start_1).IndexOf($attribute + '="') + $start_1 + $attribute.Length + 2
        $end = $html.SubString($start).IndexOf('"')
        return $html.SubString($start,$end)
    } else {
        $start = $html.SubString($start_1).IndexOf('>') + $start_1 + 1
        $end = $html.SubString($start).IndexOf('</' + $tag)
        if($end -eq -1){
            return $end
        }
        return $html.SubString($start,$end)
    }
}

function Get-TdByTh([string] $html, [string] $th){
    $th_position = $html.IndexOf($th + '</th>')
    if($th_position -eq -1){
        $th_position = $html.IndexOf($th + '</span></th>')
        if($th_position -eq -1){
            $th_position = $html.IndexOf($th + '</p></th>')
            if($th_position -eq -1){
                return $th_position
            }
        }
    }
    $start_1 = $html.SubString($th_position).IndexOf('<td') + $th_position
    $start = $html.SubString($start_1).IndexOf('>') + $start_1 + 1
    $end = $html.SubString($start).IndexOf('</td>')
    if($html.SubString($start, $end).IndexOf('<table') -ge 0){
        $end_1 = $html.SubString($start).IndexOf('</table>')
        $end = $html.SubString($start + $end_1).IndexOf('</td>') + $end_1
    }
    return $html.SubString($start, $end)
}

function Get-DataByTh([string] $html, [string] $th, [string] $tag, [string] $attribute=$null){
    $td = Get-TdByTh $html $th
    if($td -eq -1){
        return $td
    }
    return (Get-NextData $td $tag $attribute)
}

function Get-CdtId([string] $page_title){
    $page = Get-Page "" @{
        title = $page_title;
        status = 'current';
        spaceKey = $spaceName;
        expand = 'body.storage'
    }
    $html = $page.results[0].body.storage.value
    return (Get-Metadata-Value $html "cdtId")
}

function Get-Weights([string] $html){
    $ac_param_position = $html.IndexOf('V&aacute;hy zdrojov&yacute;ch kontrol, DQI (%)') 
    if($ac_param_position -eq -1){
        return $ac_param_position
    }
    $start = $html.SubString($ac_param_position).IndexOf('<table') + $ac_param_position
    $end = $html.SubString($start).IndexOf('</table>') + 8
    $table = $html.SubString($start, $end)
    $controls = @()
    for(;;){
        ## Find next title
        $start_title = $table.IndexOf('<ac:link><ri:page ri:content-title="') + 36 
        if($start_title -eq 35){
            break
        }
        $end_title = $table.SubString($start_title).IndexOf('" />')
        $title = $table.SubString($start_title, $end_title)
        
        ## Retrive dqi_key from control page
        $w_params = @{
            title = $title;
            status = 'current';
    		spaceKey = $spaceName;
    		expand = 'body.storage'            
    	}

        $s = (ConvertFrom-Json20 (Http-Get ( (Get-Url "" $w_params) ) ) )
        $dqi_id = (Get-Metadata-Value $s.results[0].body.storage.value "sourceId")
        $dqi_key = (Get-Metadata-Value $s.results[0].body.storage.value "id")
        
        ## Find next weight
        $start_weight_1 = $table.SubString($start_title + $end_title + 4).IndexOf('<td') + $start_title + $end_title + 4
        $start_weight = $table.SubString($start_weight_1).IndexOf('>') + $start_weight_1 + 1
        $end_weight = $table.SubString($start_weight).IndexOf('</td>')
        $weight = $table.SubString($start_weight, $end_weight)
        
        ## Save to array
        $controls += @{
            dqi_key=$dqi_key;
            dqi_id=$dqi_id;
            weight=(Percent-To-Double $weight)
        }
        
        ## Continue
        $table = $table.SubString($start_weight + $end_weight)       
    } 
    return $controls
}

function Get-Dimensions([string] $page_title){
    $page = Get-Page "" @{
        title = $page_title;
        status = 'current';
        spaceKey = $spaceName;
        expand = 'body.storage'
    }
    $html = $page.results[0].body.storage.value

    $table = Get-DataByTh $html 'Dimenze pravidla' 'table'
    $dimensions = @()
    if($table -eq -1){
        return $dimensions
    }
    
    for(;;){
        ## Find next title
        $start_title = $table.IndexOf('<ac:link><ri:page ri:content-title="') + 36 
        if($start_title -eq 35){
            break
        }
        $end_title = $table.SubString($start_title).IndexOf('" />')
        $title = $table.SubString($start_title, $end_title)
        
        ## Retrive dimension_key from control page
        $w_params = @{
            title = $title;
            status = 'current';
    		spaceKey = $spaceName;
    		expand = 'body.storage'            
    	}
               
        $s = (Get-Page "" $w_params)
        $dimension_id = (Get-Metadata-Value $s.results[0].body.storage.value "cdtId")
        
        ## Save to array
        $dimensions += $dimension_id
        
        ## Continue
        $table = $table.SubString($start_title + $end_title)       
    }    
    return $dimensions
}

function Percent-To-Double([string] $percent_value){
    [double]$to_double = ($percent_value.replace("%","").replace(",",".").trim())
    $output = $to_double / 100
    return $output
}

function Double-To-Percent([double] $double_value){
    if($double_value){
        return "-"
    }
    [string]$to_percent = [string]([math]::Round(($result * 100), 2)) + " %"
    $output = $to_percent.replace(".",",")
    return $output
}

function replace-str ([string] $old_string, [string] $new_string, [string] $src_path, [string] $dest_path){
    (Get-Content $src_path) | Foreach-Object {
        $_ -replace $old_string, $new_string
    } | Set-Content $dest_path
}

##
## Returns current date
##
function Get-Current-Date(){
    Get-Date -format "dd.MM.yyyy"
}

function Get-DqcDeploymentDate([string] $html){
    $deployment_date = ([string](Get-TdByTh $html "Datum nasazení"))
    
    if($deployment_date -eq "-1"){
        $deployment_date = '31.12.9999'
    } else {       
        if($deployment_date.Length -gt 10){
            if($deployment_date.IndexOf('span') -gt -1){
                $deployment_date = ([string](Get-DataByTh $html "Datum nasazení" "span"))
            } else {
                $deployment_date = '31.12.9999'
            }
        }
    }
    if($deployment_date -match '^[0-3]?[0-9]{1}\.[0-1]?[0-9]{1}\.[0-9]{4}$'){
        $datetime = [DateTime]::Parse($deployment_date)
    } else {
        $datetime = [DateTime]::Parse('31.12.9999')
    }
    return (Get-Date $datetime -format 'dd.MM.yyyy')
}

function Get-DqcData([object] $dqc){
    $html = [System.Web.HttpUtility]::HtmlDecode($dqc.body.storage.value)
    
    $cnfl_key = $dqc.id
    $key = ([string](Get-Metadata-Value $html "id")).trim()
    $id = ([string](Get-Metadata-Value $html "sourceId")).trim()
    $title = $dqc.title.Replace("(kontrola)","").trim()
    $author = ([string](Get-DataByTh $html "Autor" "ri:page" "ri:content-title")).replace("(autor)","")
    $updated_by = $dqc.version.by.displayName
    $version = $dqc.version.number
    $sql = ([string](Get-SQL-Value $html)).Replace("`n","`r`n")
    
    # Check for wrong characters
    $bytes = ([System.Text.Encoding]::ASCII.GetBytes($sql))
    
    if($bytes[0] -eq 63){
        if($bytes.length -gt 1){
            $sql = ([System.Text.Encoding]::ASCII.GetString($bytes[1..($bytes.length - 1)]))
        } else {
            $sql = ""
        }
    }
    
    $rule = ([string](Get-DataByTh $html "Kontrolovaná pravidla" "ri:page" "ri:content-title"))
    $dimensions = Get-Dimensions $rule

    $source_schema = ([string](Get-DataByTh $html "Umístění kontroly" "ri:page" "ri:content-title"))
    $source_schema_id = (Get-CdtId $source_schema)
    
    $project = ([string](Get-DataByTh $html "Projekt" "ri:page" "ri:content-title"))
    $project_id = (Get-CdtId $project)

    # $department = ([string](Get-DataByTh $html "Oddělení" "ri:page" "ri:content-title"))
    # $department_id = (Get-CdtId $department)
    
    $periodicity = ([string](Get-DataByTh $html "Periodicita" "ri:page" "ri:content-title"))
    $periodicity_id = (Get-CdtId $periodicity)
    
    $runnable = ([string](Get-Metadata-Value $html "runnable")).trim()

    # $package = ([string](Get-DataByTh $html "Workpackage" "ri:page" "ri:content-title"))
    # $package_id = (Get-CdtId $package)

    $cluster = ([string](Get-DataByTh $html "Cluster" "ri:page" "ri:content-title"))
    $cluster_id = (Get-CdtId $cluster)
    
    $deployment_date = ([string](Get-DqcDeploymentDate $html))
    
    $wave_id = (Get-CdtId $dqc.ancestors[-1].title)

    return @{
        key=$key;
        id=$id;
        cnfl_key=$cnfl_key;
        title=$title;
        author=$author;
        updated_by=$updated_by;
        version=$version;
        sql=$sql;
        source_schema=$source_schema_id;
        project=$project_id;
        # department=$department_id;
        runnable=$runnable;
        # package=$package_id;
        cluster=$cluster_id;
        deployment_date=$deployment_date;
        wave=$wave_id;
        dimensions=$dimensions;
        periodicity=$periodicity_id
    }
}

function Get-DqiData([object] $dqi){
    $html = [System.Web.HttpUtility]::HtmlDecode($dqi.body.storage.value)
    
    $key = ([string](Get-Metadata-Value $html "id")).trim()
    $id = ([string](Get-Metadata-Value $html "sourceId")).trim()
    $title = $dqi.title.Replace("(DQI)","").trim()
    $source_schema_id = -1
    $runnable = 'Y'
    $department_id = -1

    return @{
        key=$key;
        id=$id;
        title=$title;
        source_schema=$source_schema_id;
        runnable=$runnable;
        department=$department_id
    }
}

function Diff-Dqc([object] $dqc_data, [object] $db_dqc){   
    if($db_dqc){
        if($dqc_data.id -ne $db_dqc.ID){
            return 1
        }
        if((Remove-Diacritics $dqc_data.title) -ne $db_dqc.DESCR){
            return 2
        }
        if($dqc_data.source_schema -ne $db_dqc.DQ_SOURCE_SCHEMA_KEY){
            return 3
        }
        if($dqc_data.runnable -ne $db_dqc.DQ_RUNNABLE){
            return 4
        }
        if($dqc_data.periodicity -ne $db_dqc.DQ_PERIODICITY_KEY){
            return 5
        }
        if($dqc_data.project -ne $db_dqc.DQ_PROJECT_KEY){
            return 6
        }
        if($dqc_data.wave -ne $db_dqc.DQ_EG_WAVE_KEY){
            return 7
        }
        if($dqc_data.deployment_date -ne $db_dqc.DQ_DEPLOYMENT_DATE){
            return 8
        }
        if($dqc_data.cnfl_key -ne $db_dqc.CNFL_KEY){
            return 9
        }
        return 0
    }
    return -1
}

function Insert-Dqc([object] $dqc_data){
    replace-str "\[DQ_DQI_KEY\]" $dqc_data.key ("templates\DQ_DQI_Insert.sql") ("DQ_DQI.sql")
    replace-str "\[ID\]" $dqc_data.id ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DESCR\]" (Remove-Diacritics $dqc_data.title) ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_SOURCE_SCHEMA_KEY\]" $dqc_data.source_schema ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_RUNNABLE\]" $dqc_data.runnable ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_DEPARTMENT_KEY\]" $dqc_data.department ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_PROJECT_KEY\]" $dqc_data.project ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_EG_WAVE_KEY\]" $dqc_data.wave ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_DEPLOYMENT_DATE\]" $dqc_data.deployment_date ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_PERIODICITY_KEY\]" $dqc_data.periodicity ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[CNFL_KEY\]" $dqc_data.cnfl_key ("DQ_DQI.sql") ("DQ_DQI.sql")
    run-sql ("DQ_DQI.sql") 1
    Remove-Item "DQ_DQI.sql"
    $gen_key = (Get-Content "currval.txt").trim()
    Remove-Item "currval.txt"
    $dqc_data.key = $gen_key
    log ($dqc_data.id + " was inserted to RUSB_OWNER.DQ_DQI")
    return $dqc_data.key
}

function Update-Dqc([object] $dqc_data){
    replace-str "\[DQ_DQI_KEY\]" $dqc_data.key ("templates\DQ_DQI_Update.sql") ("DQ_DQI.sql")
    replace-str "\[ID\]" $dqc_data.id ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DESCR\]" (Remove-Diacritics $dqc_data.title) ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_SOURCE_SCHEMA_KEY\]" $dqc_data.source_schema ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_RUNNABLE\]" $dqc_data.runnable ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_DEPARTMENT_KEY\]" $dqc_data.department ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_PROJECT_KEY\]" $dqc_data.project ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_EG_WAVE_KEY\]" $dqc_data.wave ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_DEPLOYMENT_DATE\]" $dqc_data.deployment_date ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[DQ_PERIODICITY_KEY\]" $dqc_data.periodicity ("DQ_DQI.sql") ("DQ_DQI.sql")
    replace-str "\[CNFL_KEY\]" $dqc_data.cnfl_key ("DQ_DQI.sql") ("DQ_DQI.sql")
    run-sql ("DQ_DQI.sql") 1
    Remove-Item "DQ_DQI.sql"
    log ($dqc_data.id + " was updated in RUSB_OWNER.DQ_DQI")
}

function Insert-DqcCluster([string] $dq_clust_tp, [object] $dqc_data, [string] $dqc_cluster){
    replace-str "\[DQ_DQI_KEY\]" $dqc_data.key ("templates\DQ_DQI_CLUST_ROLE_Insert.sql") ("DQ_DQI_CLUST_ROLE.sql")
    replace-str "\[DQ_CLUST_TP_KEY\]" $dq_clust_tp ("DQ_DQI_CLUST_ROLE.sql") ("DQ_DQI_CLUST_ROLE.sql")
    replace-str "\[DQ_DATA_CLUSTER_KEY\]" $dqc_cluster ("DQ_DQI_CLUST_ROLE.sql") ("DQ_DQI_CLUST_ROLE.sql")
    run-sql ("DQ_DQI_CLUST_ROLE.sql") 1
    Remove-Item "DQ_DQI_CLUST_ROLE.sql"
    log ($dqc_data.id + " was inserted to RUSB_OWNER.DQ_DQI_CLUST_ROLE")  
}

function Update-DqcCluster([object] $cluster_role, [object] $dqc_data, [string] $dqc_cluster){
    replace-str "\[DQ_DQI_CLUST_ROLE_KEY\]" $cluster_role.DQ_DQI_CLUST_ROLE_KEY ("templates\DQ_DQI_CLUST_ROLE_Update.sql") ("DQ_DQI_CLUST_ROLE.sql")
    replace-str "\[DQ_DQI_KEY\]" $cluster_role.DQ_DQI_KEY ("DQ_DQI_CLUST_ROLE.sql") ("DQ_DQI_CLUST_ROLE.sql")
    replace-str "\[DQ_CLUST_TP_KEY\]" $cluster_role.DQ_CLUST_TP_KEY ("DQ_DQI_CLUST_ROLE.sql") ("DQ_DQI_CLUST_ROLE.sql")
    replace-str "\[DQ_DATA_CLUSTER_KEY\]" $dqc_cluster ("DQ_DQI_CLUST_ROLE.sql") ("DQ_DQI_CLUST_ROLE.sql")
    run-sql ("DQ_DQI_CLUST_ROLE.sql") 1
    Remove-Item "DQ_DQI_CLUST_ROLE.sql"
    log ($dqc_data.id + " was updated in RUSB_OWNER.DQ_DQI_CLUST_ROLE")
}

function Delete-DqcCluster([object] $cluster_role, [object] $dqc_data){
    replace-str "\[DQ_DQI_CLUST_ROLE_KEY\]" $cluster_role.DQ_DQI_CLUST_ROLE_KEY ("templates\DQ_DQI_CLUST_ROLE_Delete.sql") ("DQ_DQI_CLUST_ROLE.sql")
    run-sql ("DQ_DQI_CLUST_ROLE.sql") 1
    Remove-Item "DQ_DQI_CLUST_ROLE.sql"
    log ($dqc_data.id + " was deleted from RUSB_OWNER.DQ_DQI_CLUST_ROLE")
}

function Insert-DqcDimension([object] $dqc_data, [string] $dimension_id){
    replace-str "\[DQ_DQI_KEY\]" $dqc_data.key ("templates\DQ_DQI2DIMENSION_Insert.sql") ("DQ_DQI2DIMENSION.sql")
    replace-str "\[DQ_DIMENSION_KEY\]" $dimension_id ("DQ_DQI2DIMENSION.sql") ("DQ_DQI2DIMENSION.sql")
    run-sql ("DQ_DQI2DIMENSION.sql") 1
    Remove-Item "DQ_DQI2DIMENSION.sql"
    log ($dqc_data.id + " - Dimension " + $dimension_id + " was inserted to RUSB_OWNER.DQ_DQI2DIMENSION")  
}

function Delete-DqcDimension([object] $dqc_data, [string] $dimension_id){
    replace-str "\[DQ_DQI_KEY\]" $dqc_data.key ("templates\DQ_DQI2DIMENSION_Delete.sql") ("DQ_DQI2DIMENSION.sql")
    replace-str "\[DQ_DIMENSION_KEY\]" $dimension_id ("DQ_DQI2DIMENSION.sql") ("DQ_DQI2DIMENSION.sql")
    run-sql ("DQ_DQI2DIMENSION.sql") 1
    Remove-Item "DQ_DQI2DIMENSION.sql"
    log ($dqc_data.id + " - Dimension " + $dimension_id + " was deleted from RUSB_OWNER.DQ_DQI2DIMENSION")  
}

function Insert-DqiWeight([object] $weight, [object] $dqi_data){
    replace-str "\[parent\]" $dqi_data.key ("templates\DQ_DQI2DQI_Insert.sql") ("DQ_DQI2DQI.sql")
    replace-str "\[child\]" $weight.dqi_key ("DQ_DQI2DQI.sql") ("DQ_DQI2DQI.sql")
    replace-str "\[weight\]" $weight.weight ("DQ_DQI2DQI.sql") ("DQ_DQI2DQI.sql")
    run-sql ("DQ_DQI2DQI.sql") 1
    Remove-Item "DQ_DQI2DQI.sql"
    log ($dqi_data.id + " was inserted to RUSB_OWNER.DQ_DQI2DQI")
}

function Update-DqiWeight([object] $weight, [object] $dqi_data){
    replace-str "\[parent\]" $dqi_data.key ("templates\DQ_DQI2DQI_Update.sql") ("DQ_DQI2DQI.sql")
    replace-str "\[child\]" $weight.dqi_key ("DQ_DQI2DQI.sql") ("DQ_DQI2DQI.sql")
    replace-str "\[weight\]" $weight.weight ("DQ_DQI2DQI.sql") ("DQ_DQI2DQI.sql")
    run-sql ("DQ_DQI2DQI.sql") 1
    Remove-Item "DQ_DQI2DQI.sql"
    log ($dqi_data.id + " was updated in RUSB_OWNER.DQ_DQI2DQI")
}

function Sync-DqcCluster([int] $dq_clust_tp, [object] $dqc_data, [int] $dqc_cluster, [object] $cluster_roles){
    if($cluster_roles -and ($cluster_roles -isnot [system.array])){
        if($dqc_cluster -gt 0){
            if($cluster_roles.DQ_DATA_CLUSTER_KEY -ne $dqc_cluster){
                Update-DqcCluster $cluster_roles $dqc_data $dqc_cluster
            } else {
                log ($dqc_data.id + " - No clusters changed")
            }
        } else {
            Delete-DqcCluster $cluster_roles $dqc_data
        }
    } elseif(($dqc_cluster -gt 0) -and ($cluster_roles -isnot [system.array])){
        Insert-DqcCluster $dq_clust_tp $dqc_data $dqc_cluster
    } else {
        log ($dqc_data.id + " - No clusters changed")
    }
}

function Sync-DqcClusters([object] $dqc_data, [object[]] $cluster_role_list){
    $workpackage_roles = $cluster_role_list | Where-Object {($_.DQ_DQI_KEY -eq $dqc_data.key) -and ($_.DQ_CLUST_TP_KEY -eq $C_CDT_WORKPACKAGE_ID)}
    $cluster_roles = $cluster_role_list | Where-Object {($_.DQ_DQI_KEY -eq $dqc_data.key) -and ($_.DQ_CLUST_TP_KEY -eq $C_CDT_CLUSTER_ID)}
    
    Sync-DqcCluster $C_CDT_WORKPACKAGE_ID $dqc_data $dqc_data.package $workpackage_roles
    Sync-DqcCluster $C_CDT_CLUSTER_ID $dqc_data $dqc_data.cluster $cluster_roles
    
}

function Sync-DqcDimensions([object] $dqc_data, [object[]] $dimension_list){
    $db_final = @()
    foreach ($dim in $dqc_data.dimensions){
        $db_final += @{
            DQ_DQI_KEY=$dqc_data.key;
            DQ_DIMENSION_KEY=$dim
        }
    }
    
    $db_current = $dimension_list | Where-Object {$_.DQ_DQI_KEY -eq $dqc_data.key}
    $db_current_dimensions += $db_current | Foreach-Object {$_.DQ_DIMENSION_KEY}  
    $db_to_delete = $db_current | Where-Object {$dqc_data.dimensions -notcontains $_.DQ_DIMENSION_KEY}    
    $db_to_insert = $db_final | Where-Object {$db_current_dimensions -notcontains $_.DQ_DIMENSION_KEY}

    if($db_to_delete){
        foreach($dim in $db_to_delete){
            Delete-DqcDimension $dqc_data $dim.DQ_DIMENSION_KEY
        }
    }
    
    if($db_to_insert){
        foreach($dim in $db_to_insert){
            Insert-DqcDimension $dqc_data $dim.DQ_DIMENSION_KEY
        }
    }
}

function Create-DqcSqlFile([object] $dqc_data){
    $sqlFile =  "--------------------------------------------------" + "`r`n"
    $sqlFile += "--   DQI ID:           " + $dqc_data.id + "`r`n"
    $sqlFile += "--   Author:           " + $dqc_data.author + "`r`n"
    $sqlFile += "--   Last Updated By:  " + $dqc_data.updated_by + "`r`n"
    $sqlFile += "--   Version:          v." + $dqc_data.version + "`r`n"
    $sqlFile += "--------------------------------------------------" + "`r`n"
    # $sqlFile += "`r`n"
    $sqlFile += $dqc_data.sql
    $sqlFilePath = ".\DQI\" + $dqc_data.id + ".sql"
    # [System.IO.File]::WriteAllLines($sqlFilePath, $sqlFile, $Utf8NoBomEncoding)
    $sqlFile | Out-File -FilePath $sqlFilePath -Encoding "Default"
        
}

function Export-DQC([string] $root_page_id){
    $params = @{
        status = 'current';
        spaceKey = $spaceName;
        expand = 'body.storage,children.page,version,ancestors'
    }

    $sqlCreateRunFile =  "SET HEADING OFF FEEDBACK OFF ECHO OFF PAGESIZE 0;" + "`r`n"
    $sqlCreateRunFile += "spool run.txt;" + "`r`n"
    $sqlCreateRunFile += "select (dq_dqi_key || ':' || id || '.sql' || ':' || dq_source_schema_key)" + "`r`n"
    $sqlCreateRunFile += "from RUSB_OWNER.DQ_DQI" + "`r`n"
    $sqlCreateRunFile += "where DQ_RUNNABLE = 'Y'" + "`r`n"
    $sqlCreateRunFile += "and id in ("

    $missing = @()
    $oks = @()
    
    $dqcs = @(Get-Leaves $root_page_id $params)

    run-sql "templates\DQ_DQI_Select.sql" 1       
    $result_list = Import-Csv -Path "db_dqi.csv" -Delimiter ";" -Header "DQ_DQI_KEY","ID","DESCR","DQ_SOURCE_SCHEMA_KEY","DQ_RUNNABLE","DQ_DEPARTMENT_KEY","DQ_PROJECT_KEY","DQ_EG_WAVE_KEY","DQ_DEPLOYMENT_DATE","DQ_PERIODICITY_KEY","CNFL_KEY"
    $result_list | Out-File C:\PS_scripts\results.html
    Remove-Item "db_dqi.csv"
    
    run-sql "templates\DQ_DQI_CLUST_ROLE_Select.sql" 1
    $cluster_role_list = Import-Csv -Path "select_cluster_role.csv" -Delimiter "," -Header "DQ_DQI_CLUST_ROLE_KEY","DQ_DQI_KEY","DQ_CLUST_TP_KEY","DQ_DATA_CLUSTER_KEY"    
    Remove-Item "select_cluster_role.csv"
    
    run-sql "templates\DQ_DQI2DIMENSION_Select.sql" 1
    $dimension_list = Import-Csv -Path "select_dqi2dimension.csv" -Delimiter "," -Header "DQ_DQI_KEY","DQ_DIMENSION_KEY"  
    Remove-Item "select_dqi2dimension.csv"
    
    if($periodicity_key){
        $cnfl_keys = $result_list | Where-Object {$_.DQ_PERIODICITY_KEY -eq $periodicity_key} | Foreach {$_.CNFL_KEY}
        $cnfl_keys | Out-File C:\PS_scripts\keys.html
        $dqcs = Get-PageList $cnfl_keys
    } else {
        $dqcs = @(Get-Leaves $root_page_id $params)
    }

    $dqcCnt = 0

    foreach ($dqc in $dqcs){
    
        $dqc_data = (Get-DqcData $dqc)
 
        $dqc_result = $result_list | Where-Object {$_.DQ_DQI_KEY -eq $dqc_data.key}
        $dqc_result_id = $result_list | Where-Object {$_.ID -eq $dqc_data.id}        
        $dqc_cluster_roles = $cluster_role_list | Where-Object {$_.DQ_DQI_KEY -eq $dqc_data.key}
        $dqc_dimensions = $dimension_list | Where-Object {$_.DQ_DQI_KEY -eq $dqc_data.key}
        
        if(($dqc_data.id -ne "-1") -and ($dqc_data.sql -ne "-1") -and ($dqc_data.sql.trim() -ne "")){        
            # $has_sql += $dqc_data.id
        
            if($dqc_result -and $dqc_data.key){
                ## OK
                $oks += $dqc_data.id
                
                $diff = Diff-Dqc $dqc_data $dqc_result
                if($diff -ne 0){
                    log ($dqc_data.id  + " - Difference type: " + $diff)
                    Update-Dqc $dqc_data
                }
                Write-Host $dqc_data.id "OK"
            } elseif ($dqc_result_id){`                ## Needs revision
                ## Exists in DWH_RUSB under different DQ_DQI_KEY
                Write-Host $dqc_data.id " IN DWH_RUSB UNDER DIFFERENT KEY"
                $dqc_data.key = $dqc_result_id.DQ_DQI_KEY
                $dqc.body.storage.value = (Set-Metadata-Value $dqc.body.storage.value "id" $dqc_data.key)
                Put-Page $dqc
            } else {
                ## Not in DWH_RUSB
                Write-Host $dqc_data.id "NOT IN DWH_RUSB"                
                $missing += $dqc_data.id
                $gen_key = [string](Insert-Dqc $dqc_data)
                $dqc_data.key = $gen_key
                $dqc.body.storage.value = (Set-Metadata-Value $dqc.body.storage.value "id" $gen_key)
                Put-Page $dqc
            }
            
            if($dqcCnt -eq 0) {
                $sqlCreateRunFile += "'" + $dqc_data.id + "'"
            } else {
                $sqlCreateRunFile += ",`r`n '" + $dqc_data.id + "'"
            }
            $dqcCnt++            
            
            Sync-DqcClusters $dqc_data $cluster_role_list
            
            Sync-DqcDimensions $dqc_data $dimension_list
            
            Create-DqcSqlFile $dqc_data
            
        } else {
            ## Not Runnable
            Write-Host $dqc_data.id "MISSING SQL OR ID"
            $missing += $dqc_data.id
        }
    }

    $sqlCreateRunFile += ");" + "`r`n"
    $sqlCreateRunFile += "spool off;"

    [System.IO.File]::WriteAllLines($sqlCreateRunFilePath, $sqlCreateRunFile, $Utf8NoBomEncoding)

    log ("OK " + $oks.Length + " MISSING " + $missing.Length)
}

function Export-DQI([string] $root_page_id){
    $params = @{
        status = 'current';
        spaceKey = $spaceName;
        expand = 'body.storage,children.page,version,ancestors'
    }
    
    $dqis = @(Get-Leaves $root_page_id $params)
    
    foreach($dqi in $dqis){
        $html = $dqi.body.storage.value
        
        $dqi_data = Get-DqiData $dqi
        
        $curr_key = $dqi_data.key
        $curr_id = $dqi_data.id
        
        $weights = (Get-Weights $html)
        if($weights -and $weights -is [system.array]){
            log ($dqi.title + " has more weigths.")
            
            run-sql "templates\DQ_DQI_Select.sql" 1       
            $result_list = Import-Csv -Path "db_dqi.csv" -Delimiter ";" -Header "DQ_DQI_KEY","ID","DESCR","DQ_SOURCE_SCHEMA_KEY","DQ_RUNNABLE","DQ_DEPARTMENT_KEY","CNFL_KEY"
            Remove-Item "db_dqi.csv"
            
            run-sql "templates\DQ_DQI2DQI_Select.sql" 1       
            $dqi_to_dqi_list = Import-Csv -Path "select_dqi2dqi.csv" -Delimiter "," -Header "DQ_DQI_KEY_PARENT","DQ_DQI_KEY_CHILD","WEIGHT"
            Remove-Item "select_dqi2dqi.csv"
            
            $db_dqi = $result_list | Where-Object {$_.DQ_DQI_KEY -eq $dqi_data.key}
            $db_dqi_to_dqi = $dqi_to_dqi_list | Where-Object {$_.DQ_DQI_KEY_PARENT -eq $dqi_data.key}
            
            if(!$db_dqi){
                $gen_key = [string](Insert-Dqc $dqi_data)
                $dqi_data.key = $gen_key
                $dqi.body.storage.value = (Set-Metadata-Value $dqi.body.storage.value "id" $gen_key)
                Put-Page $dqi
            } elseif((Diff-Dqc $dqi_data $db_dqi) -ne 0){
                Update-Dqc $dqi_data
            }
            
            foreach($weight in $weights){
                $db_weight = $db_dqi_to_dqi | Where-Object {$_.DQ_DQI_KEY_CHILD -eq $weight.dqi_key}
                if($db_weight){
                    if($db_weight.WEIGHT -ne $weight.weight){
                        Update-DqiWeight $weight $dqi_data
                    }
                } else {
                    Insert-DqiWeight $weight $dqi_data
                }
            }
            continue
        }
        
        $key = $weights.dqi_key
        $id = $weights.dqi_id        
        
        if(($curr_key -eq $key) -and ($curr_id -eq $id)){
            log ($dqi.title + " was skipped.")
            continue
        }
        
        
        if(($html.IndexOf('<ac:parameter ac:name="0">sourceId</ac:parameter>') -ge 0) -and ($html.IndexOf('<ac:parameter ac:name="0">id</ac:parameter>') -ge 0)){        
            $html = Set-Metadata-Value $html "id" $key
            $html = Set-Metadata-Value $html "sourceId" $id
            $dqi.body.storage.value = $html
            Put-Page $dqi
            log ($dqi.title + " was updated.")
        } else {
            log ($dqi.title + " is missing data.")
        }
        
    }
        
}

function Init-Domain(){
    run-sql ("templates\DQ_DOMAIN_Select.sql") 1       
    $domain_list = Import-Csv -Path "select_domain.csv" -Delimiter "," -Header "DOMAIN","VALUE","DESCR"
    Remove-Item "select_domain.csv"
    
    if($domain_list){
        log ("Domain table is up to date.")
    } else {        
        log ("Domain table is missing current data.")
        run-sql ("templates\DQ_DOMAIN_Init.sql") 1
    }
}

function Sync-Codetable([string] $page_id, [string] $cdt_name){
    $codetable_data = Get-Descendants $page_id @{
        status = 'current';
        spaceKey = $spaceName;
        expand = 'body.storage,children.page,version,ancestors'
    }
    
    run-sql ("templates\" + $cdt_name + "_Select.sql") 1       
    $current_list = Import-Csv -Path "select_codetable.csv" -Delimiter "," -Header "CDT_ID","ID","DESCR"
    Remove-Item "select_codetable.csv"
    
    foreach($cdt_item in $codetable_data){
        $cdt_id = ([string](Get-Metadata-Value $cdt_item.body.storage.value "cdtId")).trim()
        $key = ([string](Get-Metadata-Value $cdt_item.body.storage.value "key")).trim()
        $value = ([string](Get-Metadata-Value $cdt_item.body.storage.value "value")).trim()
        
        if(($key) -and ($value) -and ($key -ne "") -and ($value -ne "") -and ($key -ne -1) -and ($value -ne -1)){
            if(($cdt_id -ne "-1") -and ($cdt_id -ne "")){
                $current_data = $current_list | Where-Object {$_.CDT_ID -eq $cdt_id}
                if(($current_data.ID.trim() -eq $key) -and ($current_data.DESCR.trim() -eq (Remove-Diacritics $value))){
                    # No action
                    log ("KEY: " + $cdt_id + " ID: " + $key + " DESCR: " + (Remove-Diacritics $value) + " --- SKIPPED")
                    continue
                }
                # update
                replace-str "\[cdtId\]" $cdt_id ("templates\"+$cdt_name+"_Update.sql") ($cdt_name+".sql")
                replace-str "\[key\]" $key ($cdt_name+".sql") ($cdt_name+".sql")
                replace-str "\[value\]" (Remove-Diacritics $value) ($cdt_name+".sql") ($cdt_name+".sql")
                run-sql ($cdt_name+".sql") 1
                log ("KEY: " + $cdt_id + " ID: " + $key + " DESCR: " + (Remove-Diacritics $value) + " --- UPDATED")
            } else {
                # create
                replace-str "\[key\]" $key ("templates\"+$cdt_name+"_Insert.sql") ($cdt_name+".sql")
                replace-str "\[value\]" (Remove-Diacritics $value) ($cdt_name+".sql") ($cdt_name+".sql")
                run-sql ($cdt_name+".sql") 1
                $gen_cdt_id = (Get-Content "currval.txt").trim()
                Remove-Item "currval.txt"
                $html = Set-Metadata-Value $cdt_item.body.storage.value "cdtId" $gen_cdt_id
                $page = $cdt_item
                $page.body.storage.value = $html
                Put-Page $page
                log ("KEY: " + $gen_cdt_id + " ID: " + $key + " DESCR: " + (Remove-Diacritics $value) + " --- CREATED")
            }
            Remove-Item ($cdt_name+".sql")
        }
    }
}

function Sync-ClusterToEmployee(){
    run-sql ("templates\DQ_DATA_CLUSTER_2_DQ_EMPLOYEE_Select.sql") 1       
    $current_list = Import-Csv -Path "select_codetable.csv" -Delimiter "," -Header "DQ_DATA_CLUSTER_KEY","DQ_DATA_STEWARD_KEY"
    Remove-Item "select_codetable.csv"

    $cluster_pages = Get-Descendants "94537085" @{
        status = 'current';
        spaceKey = $spaceName;
        expand = 'body.storage,children.page,version,ancestors'
    }
    
    foreach($page in $cluster_pages){
        $html = [System.Web.HttpUtility]::HtmlDecode($page.body.storage.value)
        $cluster_id = ([string](Get-Metadata-Value $html "cdtId")).trim()
        $start = $html.IndexOf('<h1>Datový stevard ČS</h1>')
        if($cluster_id -and ($cluster_id -ne -1) -and ($start -ne -1)){
            $title = Get-NextData $html.SubString($start) "ri:page" "ri:content-title"
            $steward_id = Get-CdtId $title
            $cur = $current_list | Where-Object {$_.DQ_DATA_CLUSTER_KEY -eq $cluster_id}
            if($steward_id -ne $cur.DQ_DATA_STEWARD_KEY){
                #update
                replace-str "\[employee_key\]" $steward_id ("templates\DQ_DATA_CLUSTER_2_DQ_EMPLOYEE_Update.sql") ("DQ_DATA_CLUSTER_2_DQ_EMPLOYEE_Update.sql")
                replace-str "\[cluster_key\]" $cluster_id ("DQ_DATA_CLUSTER_2_DQ_EMPLOYEE_Update.sql") ("DQ_DATA_CLUSTER_2_DQ_EMPLOYEE_Update.sql")
                run-sql ("DQ_DATA_CLUSTER_2_DQ_EMPLOYEE_Update.sql") 1
                Remove-Item ("DQ_DATA_CLUSTER_2_DQ_EMPLOYEE_Update.sql")
                log ($page.title + " --- UPDATED from " + $cur.DQ_DATA_STEWARD_KEY + " to " + $steward_id)
            } else {
                log ($page.title + " --- SKIPPED")
            }
        }
        
    }
}

function Sync-Codetables(){
    # Init-Domain

    $codetables = Import-Csv ("codetables.txt")
    
    foreach($cdt in $codetables){
        Sync-Codetable $cdt.page_id $cdt.cdt_name
    }
    
    Sync-ClusterToEmployee
}

function INIT(){
    #Sync-Codetables 
    Export-DQC $dqc_root_page
    #Export-DQI $dqi_root_page
}

############################################################################
## 
## INIT Process
##
## 0. Synchronize codetables
##
## 1. Go through all controls and check if they have record in DWH_RUSB
##    table DQ_DQI. Perform update/create accordingly. Create local SQL
##    file.
## 2. Go through all DQIs ...

if(!(Check-Connections)){
    exit
}

Load-Assemblies

INIT