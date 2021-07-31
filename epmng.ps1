# web service URL
$ws = "https://endpoints.office.com"

# assign directory and subfolders
$rootPath             = "D:\Desktop\"
$directory            = "o365_ep\"
$subfolder_Download   = "endpoints"
$subfolder_Report     = "reports"

$downloadPath = $rootPath + $directory + $subfolder_Download
$reportPath   = $rootPath + $directory + $subfolder_Report

function msgHello() {
    Write-Output "                        ============================================"
    Write-Output "                           Microsoft Office 365 Endpoints Manager   "
    Write-Output "                        ============================================"
    Write-Output ""
    Write-Output "このプログラムはMicrosoft社の公式ドキュメントの"
    Write-Output "『Office 365 URLs and IP address ranges(以下、エンドポイント)』"
    Write-Output "のデータに基づき、指定した過去のバージョンからの差分を取得するプログラムです。"
    Write-Output "エンドポイントは月に1回または2回くらい更新されています。"
    Write-Output "プログラムを実行し[Start]のメニューを選択した時点で、上記のURLの情報に変更事項がある場合"
    Write-Output "自動的に最新情報をJSONファイルとしてローカル ディレクトリに保存します"
    Write-Output "ディレクトリは規定として「D:\o365endpoints\endpoints」に設定されています。"
    Write-Output ""
    Write-Output "reference：[Office 365 URLs and IP address ranges]"
    Write-Output "           https://docs.microsoft.com/en-us/office365/enterprise/urls-and-ip-address-ranges"
    Write-Output "--------------------------------------------------------------------------------------------"
    Write-Output ""

}

#############################################
##### FUNCTION

# F1: convert decimal IP address to hex
function iptoHex([String]$inputString) {
    $marks = ".../"
    $decIp = $inputString.Split("[\.\/]")
    $str = $null

    for ($i = 0; $i -lt $decIp.Count; $i++) {
        $hexIp = [String]::Format("{0:x2}", [Int]$decIp[$i])
        $str += $hexIp
        $str += $marks[$i]
    }

    return $str

}

# F2: convert hex IP address to decimal
function iptoDec([String]$inputString) {
    $marks = ".../"
    $hexIp = $inputString.Split("[\.\/]")
    $str = $null

    for ($i = 0; $i -lt $hexIp.Count; $i++) {
        $decIp = [Convert]::ToInt32($hexIp[$i], 16)
        $str += [String]$decIp
        $str += $marks[$i]
    }

    return $str

}

# F3: check for duplication
function dpCheck([String]$inputString, [System.Array]$inputArray) {
    return $inputArray -contains $inputString

}


##########################################################################
##########################################################################

msgHello
Write-Output "0. Exit"
Write-Output "1. Start"

$inputValue = $null
do {
    [String]$inputValue = Read-Host "Input"

} while (!($inputValue -eq 0 -or $inputValue -eq 1))

switch ($inputValue) {
    0 { exit }
    1 {
        $clientRequestId = [GUID]::NewGuid().Guid
        $deviceVersion = "current_version.txt"

        ## ローカルにGUIDおよび最新バージョンの情報があるか確認
        if (Test-Path $rootPath$directory$deviceVersion) {
            $currentVersion = Get-Content $rootPath$directory$deviceVersion

        } else {
            switch (Test-Path $rootPath$directory) { 
                $False { New-Item -Path $rootPath -Name $directory -ItemType "directory" } 
                $True { }
            }

            $currentVersion = "0000000000"
            $currentVersion | Out-File $rootPath$directory$deviceVersion

        }

        ## 更新されたendpointsのバージョンがあるか確認
        $version = Invoke-RestMethod -Uri ($ws + "/version/Worldwide?clientRequestId=" + $clientRequestId)
        if ($version.latest -gt $currentVersion) {
            Write-Output ""
            Write-Output "========================================================================================"
            Write-Output "   New version of Office 365 worldwide commercial service instance endpoints detected   "
            Write-Output "========================================================================================"
            Write-Output ""

            $webclient = New-Object System.Net.WebClient

            Write-Output ""
            Write-Output "Downloading..."
            Write-Output ""

            ## $downloadPath(D:\o365endpoints\endpoints)が存在しない場合、新しいディレクトリを作成
            switch (Test-Path $downloadPath) {
                $True { Break }
                $False { New-Item -Path $rootPath$directory -Name $subfolder_Download -ItemType "directory" }

            }

            ## download the latest version of endpoints (noipv6)
            $url = "$ws/endpoints/Worldwide?noipv6&clientRequestId=$clientRequestId"
            $file = $downloadPath + $version.latest + ".json"
            $webclient.DownloadFile($url, $file)

            ## download the history of endpoints (noipv6)
            $url = "$ws/changes/worldwide/0000000000?noipv6&clientRequestId=$clientRequestId"
            $file = $rootPath + $directory + "history.json"
            $webclient.DownloadFile($url, $file)

            ## 更新された$version.latestの情報をローカルに保存 
            $version.latest | Out-File $rootPath$directory$currentVersion

            Write-Output ""
            Write-Output "============================================================================="
            Write-Output "  Office 365 worldwide commercial service instance endpoints are up-to-date  "
            Write-Output "============================================================================="
            Write-Output ""

        }
        else {
            Write-Output ""
            Write-Output "============================================================================="
            Write-Output "  Office 365 worldwide commercial service instance endpoints are up-to-date  "
            Write-Output "============================================================================="
            Write-Output ""

        } 

        Write-Output "0. Exit"
        Write-Output "1. Write a report"

        $inputValue = $null
        do {
            [String]$inputValue = Read-Host "Input"

        } while (!($inputValue -eq 0 -or $inputValue -eq 1))

        Write-Output ""
        switch ($inputValue) {
            0 { exit }
            1 {
                $hPath = $rootPath + $directory + "history.json"
                $history = Get-Content -Path $hPath | ConvertFrom-Json
                $versionHistory = $history | Select-Object -Property "version" -Unique | Sort-Object "version" -Descending

                Write-Output "select a version to compare with"
                Write-Output " 0. Exit"
                
                $refNum = 0
                for ($i = 1; $i -lt $versionHistory.Count; $i++) {
                    $str = $null
                    $file = $downloadPath + $versionHistory[$i].version + ".json"

                    if (Test-Path $file) {
                        $refNum += 1

                        switch ($true) {
                            ($refNum -lt 10) { $str = " " + "$refNum" + '. ' + $versionHistory[$i].version }
                            ## ($refNum -ge 10) { $str = "$refNum" + '. ' + $versionHistory[$i].version }
                            ($refNum -eq 10) { $str = " a. see more" }

                        }

                    }
                    Write-Output $str

                }

            }

        }

        ######TODO#####
        ## 日付を指定して変数に保存
        if ($refNum -lt 10) {
            $iDate = $null
            do {
                $iDate = Read-Host "Input"
    
                if ($iDate -eq 0) { exit }
                else { }
    
            } while (!($iDate -match "^[1-9]+[0-9]*$" -and [Int]$iDate -ge 1 -and [Int]$iDate -le $refNum))

        } else {
            $refNum = 9
            $aSwitch = $False

            $iDate = $null
            do {
                $iDate = Read-Host "Input"
    
                switch ($iDate) {
                    0 { exit }
                    a {
                        if ($aSwitch -eq $False) {
                            $aSwitch = $True
                            Write-Output ""
                            Write-Output "＜バージョン一覧＞"
                            Write-Output " 0. Exit"
                            $refNum = 0
                            for ($i = 1; $i -lt $versionHistory.Count; $i++) { 
                                $str = $null 
                                $file = $downloadPath + $versionHistory[$i].version + ".json"
                    
                                if (Test-Path $file) {
                                    $refNum += 1
                    
                                    switch ($true) {
                                        ($refNum -lt 10) { $str = " " + "$refNum" + '. ' + $versionHistory[$i].version }
                                        ($refNum -ge 10) { $str = "$refNum" + '. ' + $versionHistory[$i].version }
        
                                    }
                                    Write-Output $str
                    
                                }
                            }
                        }
                    }
    
                }
    
            } while (!($iDate -match "^[1-9]+[0-9]*$" -and [Int]$iDate -ge 1 -and [Int]$iDate -le $refNum))

        }

        ## 最新版endpointsを呼び出す
        $fLatest = $downloadPath + $version.latest + ".json"
        $eLatest = Get-Content -Path $fLatest | ConvertFrom-Json

        ## 指定した日付のendpointsを呼び出す
        $fIndicated = $downloadPath + $versionHistory[$iDate].version + ".json"
        $eIndicated = Get-Content -Path $fIndicated | ConvertFrom-Json

        ## 比較の結果を保存するテーブルを宣言
        $aList = New-Object system.Data.DataTable
        $aList.Columns.Add("no")
        $aList.Columns.Add("change")
        $aList.Columns.Add("isDeleted")
        $aList.Columns.Add("type")
        $aList.Columns.Add("id")
        $aList.Columns.Add("serviceArea")
        $aList.Columns.Add("serviceAreaDisplayName")
        $aList.Columns.Add("key")
        $aList.Columns.Add("tcpPorts")
        $aList.Columns.Add("udpPorts")
        $aList.Columns.Add("expressRoute")
        $aList.Columns.Add("category")
        $aList.Columns.Add("required")
        $aList.Columns.Add("notes")
 
        ## 比較に必要なIDの集合 
        $ids = $eLatest.id + $eIndicated.id | Sort-Object -Unique

        ## サブプロパティ(URL, IPアドレス以外の属性)を保存するテーブルを作成
        $pList = New-Object system.Data.DataTable
        $pList.Columns.Add("id") 
        $pList.Columns.Add("serviceArea") 
        $pList.Columns.Add("serviceAreaDisplayName") 
        $pList.Columns.Add("tcpPorts") 
        $pList.Columns.Add("udpPorts") 
        $pList.Columns.Add("expressRoute") 
        $pList.Columns.Add("category") 
        $pList.Columns.Add("required") 
        $pList.Columns.Add("notes") 

        for ($i = 0; $i -lt $ids.Count; $i++) { 
            if ($eLatest.id -contains $ids[$i]) { $subProperty = $eLatest | Where-Object -Property id -EQ -Value $ids[$i] } 
            else { $subProperty = $eIndicated | Where-Object -Property id -EQ -Value $ids[$i] } 

            $pList.Rows.Add($subProperty.id, $subProperty.serviceArea, $subProperty.serviceAreaDisplayName 
                , $subProperty.tcpPorts, $subProperty.udpPorts, $subProperty.expressRoute 
                , $subProperty.category, $subProperty.required, $subProperty.notes) 
        } 

        ## 比較の結果をテーブルに記録 
        $no = 0 

        for ($i = 0; $i -lt $ids.Count; $i++) { 
            # 最新バージョンのIDに該当するオブジェクト 
            $ref = $eLatest | Where-Object -Property id -EQ -Value $ids[$i] 

            if ($null -eq $ref) { 
                $ref = @() 
                Add-Member -InputObject $ref -MemberType NoteProperty -Name urls -Value @() 
                Add-Member -InputObject $ref -MemberType NoteProperty -Name ips -Value @() 
            } 

            # 指定した日付のIDに該当するオブジェクト 
            $dif = $eIndicated | Where-Object -Property id -EQ -Value $ids[$i] 

            if ($null -eq $dif) {
                $dif = @()
                Add-Member -InputObject $dif -MemberType NoteProperty -Name urls -Value @()
                Add-Member -InputObject $dif -MemberType NoteProperty -Name ips -Value @()
            } 

            $prop = $pList | Where-Object -Property id -EQ -Value $ids[$i]
            $inCategory = $eLatest | Where-Object -Property serviceArea -EQ -Value $prop.serviceArea

            switch ($True) {
                ($null -eq $ref.urls) { Add-Member -InputObject $ref -MemberType NoteProperty -Name urls -Value @() }
                ($null -eq $ref.ips) { Add-Member -InputObject $ref -MemberType NoteProperty -Name ips -Value @() }
                ($null -eq $dif.urls) { Add-Member -InputObject $dif -MemberType NoteProperty -Name urls -Value @() }
                ($null -eq $dif.ips) { Add-Member -InputObject $dif -MemberType NoteProperty -Name ips -Value @() }
            } 

            $sUrl = Compare-Object -ReferenceObject @($ref.urls) -DifferenceObject @($dif.urls) -IncludeEqual | Sort-Object -Property InputObject 
            $sIp = Compare-Object -ReferenceObject @($ref.ips) -DifferenceObject @($dif.ips) -IncludeEqual 
            $sIpHex = @($sIp) | ForEach-Object { iptoHex $_.InputObject } | Sort-Object 
            $sIpDec = @($sIpHex) | ForEach-Object { iptoDec $_ } 
            $sIpIndicator = New-Object System.Collections.ArrayList 

            for ($j = 0; $j -lt $sIpDec.Count; $j++) { 
                for ($k = 0; $k -lt $sIp.Count; $k++) { 
                    if (@($sIpDec)[$j] -eq @($sIp)[$k].InputObject) { $sIpIndicator.Add(@($sIp)[$k].SideIndicator) } 

                }
            } 

            switch ($sUrl.Count) { 
                $null { $uCount = 1 } 
                default { $uCount = $sUrl.Count } 
            } 

            for ($j = 0; $j -lt $uCount; $j++) { 
                $no += 1 
                switch (@($sUrl)[$j].sideIndicator) { 
                    "<=" {
                        $impact = "Added"
                        $dpChecked = ""
                    } 
                    "==" {
                        $impact = ""
                        $dpChecked = ""
                    }
                    "=>" { 
                        ## カテゴリ(serviceArea)内での重複を検査 
                        $check = dpCheck @($sUrl)[$j].InputObject $inCategory.urls 
                        if ($check -eq $True) { $impact = "RemovedDuplicateIpOrUrl" } 
                        else { $impact = "Removed" } 

                        ## すべてのendpointでの重複を検査 
                        $check = dpCheck @($sUrl)[$j].InputObject $eLatest.urls 
                        if ($check -eq $True) { $dpChecked = "" } 
                        else { $dpChecked = "DELETED" } 

                    } 
                } 
            
                $aList.Rows.Add($no, $impact, $dpChecked, "URL", $prop.id, $prop.serviceArea, $prop.serviceAreaDisplayName 
                    , @($sUrl)[$j].InputObject, $prop.tcpPorts, $prop.udpPorts, $prop.expressRoute 
                    , $prop.category, $prop.required, $prop.notes) 

            } 

            switch ($sIp.Count) { 
                $null { $iCount = 1 } 
                default { $iCount = $sIp.Count } 
            } 

            for ($j = 0; $j -lt $iCount; $j++) { 
                $no += 1 

                switch (@($sIpIndicator)[$j]) { 
                    "<=" {
                        $impact = "Added"
                        $dpChecked = ""
                    }
                    "==" {
                        $impact = ""
                        $dpChecked = ""
                    }
                    "=>" { 
                        ## カテゴリ(serviceArea)内での重複を検査 
                        $check = dpCheck @($sIp)[$j] $inCategory.ips 
                        if ($check -eq $True) { $impact = "RemovedDuplicateIpOrUrl" } 
                        else { $impact = "Removed" } 

                        ## すべてのendpointでの重複を検査 
                        $check = dpCheck @($sIp)[$j] $eLatest.ips 

                        if ($check -eq $True) { $dpChecked = "" } 
                        else { $dpChecked = "DELETED" } 

                    } 
                } 

                $aList.Rows.Add($no, $impact, $dpChecked, "IP", $prop.id, $prop.serviceArea, $prop.serviceAreaDisplayName 
                    , @($sIpDec)[$j], $prop.tcpPorts, $prop.udpPorts, $prop.expressRoute 
                    , $prop.category, $prop.required, $prop.notes) 

            } 

        } 

    }

}

## 比較した結果を保存するパスを指定 
if (!(Test-Path $reportPath)) { New-Item -Path $rootPath$directory -Name $subfolder_Report -ItemType "directory" }

$time = Get-Date -Format "yyyyMMdd_HHmmss"
$fileName = $subfolder_Report.Substring(0, $subfolder_Report.Length - 1) + "($time).csv"
$aList | ConvertTo-Csv -NoTypeInformation -UseCulture | Out-File -FilePath $reportPath$fileName

# Define locations and delimiter
$csv = "$reportPath\report($time).csv" #Location of the source file
$xlsx = "$reportPath\report($time).xlsx" #Desired location of output
$delimiter = "," #Specify the delimiter used in the file

# Create a new Excel workbook
$excel = New-Object -ComObject excel.application
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)

# Build the QueryTables.Add command and reformat the data
$txtConnector = ("TEXT;" + $csv)
$connector = $worksheet.QueryTables.add($txtConnector, $worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType = 1
$query.TextFileColumnDataTypes = , 1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

$excel.Rows.Item("2:2").Select()
$excel.ActiveWindow.FreezePanes = $True

$worksheet.Range("A1:N1").interior.colorindex = 20

$query.Refresh()
$query.Delete()

$workbook.SaveAs($xlsx, 51)
$excel.Quit()