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
# params

$spaceName = $h.spaceName
$username = $h.username
$password = $h.password

$dqi_root_page = $h.dqi_root_page

$rest_api_url = $h.rest_api_url

###########################################################################
# sql client path

$sqlplus = '"' + $h.sqlplus + '"'

###########################################################################
# Define Web client

$web = New-Object Net.WebClient

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
    $sql_log_file = "logs\" + (Get-Date -format "yyyy-MM-dd") + "_publish.log"
    $log_message = (Get-Date -format "yyyy-MM-dd HH:mm:ss") + " " + $message
    Add-Content $sql_log_file $log_message
    Write-Host $log_message
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

function get-connection ([int] $schema_id){
    foreach($con in $connections){
        if($con.schema_id -eq $schema_id){
            return $con
        }
    }
    return $null
}

function run-sql ([string] $path, [int] $source_schema_id) {
    log ("Running " + $path)
    $sql_log_file = "logs\" + (Get-Date -format "yyyy-MM-dd") + "_publish.log"
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

function Get-Metadata-Value([string] $html, [string] $name){
    $ac_param_position = $html.IndexOf('<ac:parameter ac:name="0">' + $name + '</ac:parameter>') 
    if($ac_param_position -eq -1){
        return $ac_param_position
    }
    $start = $html.SubString($ac_param_position).IndexOf('<![CDATA[') + $ac_param_position + 9
    $end_macro = $html.SubString($ac_param_position).IndexOf('</ac:structured-macro>') + $ac_param_position
    if(($end_macro + 9) -lt $start){
        return ""
    }
    $end = $html.SubString($start).IndexOf(']]')
    return $html.SubString($start, $end)
}

function Percent-To-Double([string] $percent_value){
    [double]$to_double = ($percent_value.replace("%","").replace(",",".").trim())
    $output = $to_double / 100
    return $output
}

function Double-To-Percent([double] $double_value, [int] $decimal_digits = 2){
    if($double_value -lt 0){
        return "-"
    }
    [string]$to_percent = [string]([math]::Round(($result * 100), $decimal_digits)) + " %"
    $output = $to_percent.replace(".",",")
    return $output
}

function Get-Current-Date(){
    Get-Date -format "dd.MM.yyyy"
}

function Out-Excel([string] $Path = "$env:temp\$(Get-Date -Format yyyyMMddHHmmss).csv") {
  $input | Export-CSV -Path $Path -UseCulture -Encoding UTF8 -NoTypeInformation
  Invoke-Item -Path $Path
}

function Export-ToCsv([string] $page_id, [object] $params){
    $pages = Get-Leaves $page_id $params    
    $page_obj = @()
    
    # New-Item -ItemType file $output –force
    foreach($page in $pages){
        $obj = New-Object System.Object
        $obj | Add-Member -type NoteProperty -name Id -value $page.id
        $obj | Add-Member -type NoteProperty -name Title -value $page.title
        $obj | Add-Member -type NoteProperty -name Content -value $page.body.storage.value
        $page_obj += $obj        
    }
    $page_obj | Out-Excel "CNFL_Export_$(Get-Date -Format yyyyMMddHHmmss).csv"
}

function Import-FromCsv([string] $Path){
    if(!$Path){
        return
    }
    
    $i_params = @{
        status = 'current';
		spaceKey = $spaceName;
		expand = 'body.storage,children.page,version,ancestors'
	}
    
    Get-Content $Path | Set-Content -Encoding utf8 "utf8_encoded_import.csv"
    $pages = Import-CSV -Path "utf8_encoded_import.csv" -Delimiter "`t"
    Remove-Item "utf8_encoded_import.csv"
    
    Write-Host $pages
    
    $pages | Foreach-Object {        
        $page = (Get-Page $_.Id $i_params)
        
        $ancestors = @()
        $ancestors += @{
            id=$page.ancestors[-1].id;
            type="page"
        }

        $page.title = $_.Title
        $page.body.storage.value = $_.Content
        $page.ancestors = $ancestors
        $page.version.number = ($page.version.number + 1)        
        
        $json = ConvertTo-Json20 $page
        
        Http-Put (Get-Url $_.Id) $json
    }
}

$params = @{
    status="current";
    spaceKey=$spaceName;
    expand="body.storage,children.page,version,ancestors"
    # limit="100"
}

############################################################################
## DQI Specific

function Get-ThYellow ([string] $html){
    return (Percent-To-Double (Get-Metadata-Value $html "tresholdYellow"))
}

function Get-ThGreen ([string] $html){  
    return (Percent-To-Double (Get-Metadata-Value $html "tresholdGreen")) 
}

function Get-Weights([string] $html){
    $ac_param_position = $html.IndexOf('<ac:parameter ac:name="0">tresholdGreen</ac:parameter>') 
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

function Get-Tresholds([string] $html){
    $ac_param_position = $html.IndexOf('Prahy kvality') 
    if($ac_param_position -eq -1){
        return $ac_param_position
    }
    $start = $html.SubString($ac_param_position).IndexOf('<table') + $ac_param_position
    $end = $html.SubString($start).IndexOf('</table>') + 8
    $table = $html.SubString($start, $end)
    $tresholds = @()
    for(;;){
        ## Find next title
        $start_title = $table.IndexOf('<ac:link><ri:page ri:content-title="') + 36 
        if($start_title -eq 35){
            break
        }
        $end_title = $table.SubString($start_title).IndexOf('" />')
        $title = $table.SubString($start_title, $end_title)
        
        ## Find avg treshold
        $start_at_1 = $table.SubString($start_title + $end_title + 4).IndexOf('<td') + $start_title + $end_title + 4
        $start_at = $table.SubString($start_at_1).IndexOf('>') + $start_at_1 + 1
        $end_at = $table.SubString($start_at).IndexOf('</td>')
        $avg_treshold = $table.SubString($start_at, $end_at)
        
        ## Find sat treshold
        $start_st_1 = $table.SubString($start_at + $end_at + 4).IndexOf('<td') + $start_at + $end_at + 4
        $start_st = $table.SubString($start_st_1).IndexOf('>') + $start_st_1 + 1
        $end_st = $table.SubString($start_st).IndexOf('</td>')
        $sat_treshold = $table.SubString($start_st, $end_st)
        
        ## Save to array
        $tresholds += @{
            title = $title;
            avg=(Percent-To-Double ([System.Web.HttpUtility]::HtmlDecode($avg_treshold)));
            sat=(Percent-To-Double ([System.Web.HttpUtility]::HtmlDecode($sat_treshold)));
            digits=(Get-DecimalDigits ([System.Web.HttpUtility]::HtmlDecode($avg_treshold)) ([System.Web.HttpUtility]::HtmlDecode($sat_treshold)))
        }
        
        ## Continue
        $table = $table.SubString($start_st + $end_st)       
    } 
    return $tresholds
}

# function Get-Image ([string] $html_current, [double] $result = $null){
function Get-Image ([double] $th_yellow, [double] $th_green, [double] $result = $null){
    # $th_yellow = (Get-ThYellow $html_current)
    # $th_green = (Get-ThGreen $html_current)
    
    if($result -lt 0){
        return '<ac:image ac:thumbnail="true" ac:width="66"><ri:attachment ri:filename="IndicatorInactive.jpg"><ri:page ri:content-title="Semafory DQI" /></ri:attachment></ac:image>' 
    } elseif (($result -ge $th_green) -and $th_green) {
        return '<ac:image ac:thumbnail="true" ac:width="66"><ri:attachment ri:filename="IndicatorGreen.jpg"><ri:page ri:content-title="Semafory DQI" /></ri:attachment></ac:image>'
    } elseif (($result -ge $th_yellow) -and $th_yellow){
        return '<ac:image ac:thumbnail="true" ac:width="66"><ri:attachment ri:filename="IndicatorYellow.jpg"><ri:page ri:content-title="Semafory DQI" /></ri:attachment></ac:image>'
    } else {
        return '<ac:image ac:thumbnail="true" ac:width="66"><ri:attachment ri:filename="IndicatorRed.jpg"><ri:page ri:content-title="Semafory DQI" /></ri:attachment></ac:image>' 
    }
    
}

function Get-DecimalDigits([string] $html_current){
    [string]$th_yellow = (Get-ThYellow $html_current)
    [string]$th_green = (Get-ThGreen $html_current)
    if($th_yellow.IndexOf(".") -ge 0){
        $th_yellow_digits = $th_yellow.Split(".")[1].Trim().Length - 2
    } else {
        $th_yellow_digits = 2
    }
    if($th_green.IndexOf(".") -ge 0){
        $th_green_digits = $th_green.Split(".")[1].Trim().Length - 2
    } else {
        $th_green_digits = 2
    }
    return (($th_yellow_digits,$th_green_digits,2 | Measure -Max).Maximum)
}

##
## Returns HTML which should be assigned to [JSON].body.storage.value
##
function Set-Result ([string] $html_current, [double] $result = -1, [object] $treshold){  
    [string]$result_to_display = (Double-To-Percent $result $treshold.digits)
    $html = $html_current
    
    $insert = '<tr><th style="text-align: right;">Hodnota DQI - <ac:link><ri:page ri:content-title="'
    $insert += $treshold.title
    $insert += '" /></ac:link></th><td><h1 style="text-align: center;">'
    $insert += $result_to_display
    $insert += '</h1></td><td>'
    $insert += (Get-Image $treshold.avg $treshold.sat $result)
    $insert += '</td></tr>'
    
    # $ac_param_position = $html.IndexOf('<ac:parameter ac:name="0">tresholdGreen</ac:parameter>') 
    
    # Write-Host $html
    
    $ac_param_position = $html.IndexOf('Hodnota DQI - <ac:link><ri:page ri:content-title="' + $treshold.title + '" /></ac:link>') 
    # Write-Host $ac_param_position
    if($ac_param_position -eq -1){
        $start_date_1 = $html.IndexOf('Datum aktualizace</th>')
        if($start_date_1 -eq -1){
            log("Na stránke sa nenachádza Datum aktualizace")
            return $html_current
        }
        $start_pos = $html.SubString($start_date_1 - 60).IndexOf('<tr>') + $start_date_1 - 60
        $end_pos = $start_pos
    } else {
        $start_pos = $html.SubString($ac_param_position - 50).IndexOf('<tr>') + $ac_param_position - 50
        $end_pos = $html.SubString($start_pos + 4).IndexOf('<tr>') + $start_pos + 4
    }
    
    $html = $html.SubString(0,$start_pos) + $insert + $html.SubString($end_pos)
    
    <#
    ## Set results
    $start_result_1 = $html.SubString($ac_param_position).IndexOf('<h1') + $ac_param_position
    $start_result = $html.SubString($start_result_1).IndexOf('>') + $start_result_1 + 1
    $end_result = $html.SubString($start_result).IndexOf('<') + $start_result
    $html = $html.SubString(0,$start_result) + $result_to_display + $html.SubString($end_result)
    
    ## Update indicator image
    $start_img = $html.SubString($start_result).IndexOf('<ac:image') + $start_result
    if($start_img -eq -1){
        return $html_current
    }
    $end_img = $html.SubString($start_img).IndexOf('</ac:image>') + $start_img + 11
    $html = $html.SubString(0,$start_img) + (Get-Image $treshold.avg $treshold.sat $result) + $html.SubString($end_img)
    #>
    
    return $html
}

function Update-Date([string] $html_current, [string] $bus_date = ""){
    $html = $html_current

    ## Update date of change
    $start_date_1 = $html.IndexOf('Datum aktualizace</th>')
    $start_date_2 = $html.SubString($start_date_1).IndexOf('<td') + $start_date_1
    $start_date = $html.SubString($start_date_2).IndexOf('>') + $start_date_2 + 1
    $end_date = $html.SubString($start_date).IndexOf('</td>') + $start_date    
    $html = $html.SubString(0,$start_date) + (Get-Current-Date) + $html.SubString($end_date)
    
    ## Update business date
    $start_bus_date_1 = $html.SubString($start_date).IndexOf('Rozhodné datum</th>') + $start_date
    $start_bus_date_2 = $html.SubString($start_bus_date_1).IndexOf('<td') + $start_bus_date_1
    $start_bus_date = $html.SubString($start_bus_date_2).IndexOf('>') + $start_bus_date_2 + 1
    $end_bus_date = $html.SubString($start_bus_date).IndexOf('</td>') + $start_bus_date    
    $html = $html.SubString(0,$start_bus_date) + $bus_date + $html.SubString($end_bus_date)
    
    return $html
}

function Dqi-UpdateResult([string] $page_id, [object] $params, [double] $result = -1, [double] $vol_result = -1, [string] $bus_date = ""){
    $dqi = (Get-Page $page_id $params)
    $html = $dqi.body.storage.value
    $tresholds = Get-Tresholds $html
    $th_count = $tresholds | Where-Object {$_.title -eq 'Count based (kvantifikace)'}
    $th_volume = $tresholds | Where-Object {$_.title -eq 'Volume based (kvantifikace)'}
    
    if($th_count){
        $html = [string](Set-Result $html $result $th_count)
    }
    if($th_volume){
        $html = [string](Set-Result $html $vol_result $th_volume)
    }
    
    $html = [string](Update-Date $html $bus_date)
    
    $dqi.body.storage.value = $html  
    # Write-Host $html
    [string]$output = Put-Page $dqi
    if($output){
        log($dqi.title + " was updated.")
    } else {
        log($dqi.title + " failed to be updated due to: " + $output)
    }
}

function Print-Leaves([string] $page_id, [object] $params){
    $dqis = Get-Leaves $page_id $params

    foreach($dqi in $dqis){
        Write-Host $dqi.id $dqi.title
    }
}

function Dqi-GetRate ([string] $dqi_id, [object[]] $results){
    foreach($res in $results){
        if($res.id.trim() -eq $dqi_id){
            return $res.rate
        }
    }
    return -100
}

##
## Updates all EG DQIs
##
function Dqi-Update(){
    ## run-sql ".\templates\results.sql" 1
    run-sql ".\templates\DQ_DQI_RESULT_Select.sql" 1
    ## $results = Import-Csv -Path "result_output.csv" -header "key","rate","bus_date"
    $results = Import-Csv -Path "result_output.csv" -header "key","wave","periodicity","rate","vol_rate","bus_date"
    Remove-Item "result_output.csv"
    
    if($periodicity_key){
        $results = ($results | Where-Object {$_.periodicity -eq $periodicity_key})
        $root = "70250069"
    } else {
        $root = $dqi_root_page
    }
    
    ## $bus_date = $results[0].bus_date
    
    $params = @{
        status="current";
        spaceKey=$spaceName;
        expand="body.storage,children.page,version,ancestors"
    }
    
    $dqis = Get-Leaves $root $params
    
    foreach($dqi in $dqis){
        $key = Get-Metadata-Value $dqi.body.storage.value "id"
        $result = ($results | Where-Object {$_.key -eq $key})
        
        if($result){    
            if($result.rate -or $result.vol_rate){
                Dqi-UpdateResult $dqi.id $params $result.rate $result.vol_rate $result.bus_date 
            } else {
                Dqi-UpdateResult $dqi.id $params
            }
        }
        
    }        
}

function Set-Metadata-Value([string] $html, [string] $name, [string] $value){
    $ac_param_position = $html.IndexOf('<ac:parameter ac:name="0">' + $name + '</ac:parameter>') 
    if($ac_param_position -eq -1){
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

###########################################################################
# PROCESS

Load-Assemblies

Dqi-Update