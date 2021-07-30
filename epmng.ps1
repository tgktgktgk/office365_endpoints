## 関数：10進数 -> 16進数
function toHex([String]$inputString) {
    $marks = ".../"
    $decNum = $inputString.Split("[\.\/]")
    $str = $null

    for ($i = 0; $i -lt $decNum.Count; $i++) {
        $hexNum = [String]::Format("{0:x2}", [Int]$decNum[$i])
        $str += $hexNum
        $str += $marks[$i]
    }
    return $str 

}

## 関数：16進数 -> 10進数
function toDec([String]$inputString) {
    $marks = ".../"
    $hexNum = $inputString.Split("[\.\/]")
    $str = $null

    for ($i = 0; $i -lt $hexNum.Count; $i++) {
        $decNum = [Convert]::ToInt32($hexNum[$i], 16)
        $str += [String]$decNum
        $str += $marks[$i]
    }
    return $str

}

## メッセージ：プログラム起動
function msgHello() {
    Write-Output "---------------------------------"
    Write-Output "| Managing Office 365 endpoints |"
    Write-Output "---------------------------------"
    Write-Output ""
    Write-Output "このプログラムはMicrosoft社の公式ドキュメントの"
    Write-Output "『Office 365 URLs and IP address ranges(以下、エンドポイント)』"
    Write-Output "のデータに基づき、指定した過去のバージョンからの差分を取得するプログラムです。"
    Write-Output "エンドポイントは月に1回または2回くらい更新されています。"
    Write-Output "プログラムを実行し[Start]のメニューを選択した時点で、上記のURLの情報に変更事項がある場合"
    Write-Output "自動的に最新情報をJSONファイルとしてローカル ディレクトリに保存します"
    Write-Output "ディレクトリは規定として「D:\o365endpoints\endpoints」に設定されています。"
    Write-Output ""
    Write-Output "参考：https://docs.microsoft.com/en-us/office365/enterprise/urls-and-ip-address-ranges"
    Write-Output "-----------------------------------------------------------------------------------------"
    Write-Output ""
}

## メッセージ：アップデートの検知 
function msgDetected() {
    Write-Output ""
    Write-Output "======================================================================================="
    Write-Output "New version of Office 365 worldwide commercial service instance endpoints detected"
    Write-Output "======================================================================================="
    Write-Output "" 
}

## メッセージ：アップデートの完了 
function msgCompleted() {
    Write-Output ""
    Write-Output "============================================================================"
    Write-Output "Office 365 worldwide commercial service instance endpoints are up-to-date"
    Write-Output "============================================================================"
    Write-Output ""
}

## 関数：指定した範囲内で重複の検査
function dpCheck([String]$inputString, [System.Array]$inputArray) {
    return $inputArray -contains $inputString
}

##########
##########

## Webサービス URL 
$ws = "https://endpoints.office.com" 

## GUIDおよび最新バージョンの日付を保存するパスを指定 
$datapath = "D:\" 
$foldername = "o365endpoints\" 
$filename = "endpoints_clientid_latestversion.txt" 

## endpointsが保存されているパスを指定 
$savepath = $datapath + $foldername + "endpoints\" 

<#
$datapath$foldername$filename : "D:\o365endpoints\endpoints_clientid_lastversion.txt"
#> 

msgHello
Write-Output "0. Exit"
Write-Output "1. Start"

$direction = $null 
do { 
    [String]$direction = Read-Host "入力"

} while (!($direction -eq 0 -or $direction -eq 1))

switch ($direction) {
    0 { exit } #switch ($direction) condition (0) end //

    1 {
        ## ローカルにGUIDおよび最新バージョンの情報があるか確認
        if (Test-Path $datapath$foldername$filename) {
            $content = Get-Content $datapath$foldername$filename
            $clientRequestId = $content[0]
            $latestVersion = $content[1]

        } else { 
            ## $datapath$foldername(D:\o365endpoints\)が存在しない場合、新しいディレクトリを作成 
            switch (Test-Path $datapath$foldername) { 
                $False { New-Item -Path $datapath -Name $foldername -ItemType "directory" } 
                $True { }
            } 

            ## GUIDの発行およびlatestVersionの初期化 
            $clientRequestId = [GUID]::NewGuid().Guid 
            $latestVersion = "0000000000" 
            @($clientRequestId, $latestVersion) | Out-File $datapath$foldername$filename

        }

        ## 更新されたendpointsのバージョンがあるか確認
        $version = Invoke-RestMethod -Uri ($ws + "/version/Worldwide?clientRequestId=" + $clientRequestId)
        if ($version.latest -gt $latestVersion) {
            ## メッセージ：アップデートの検知
            msgDetected

            ## jsonファイルをダウンロードするために必要なwebclientを宣言 
            $webclient = New-Object System.Net.WebClient

            Write-Output ""
            Write-Output "ダウンロード中"
            Write-Output ""

            ## $savepath(D:\o365endpoints\endpoints)が存在しない場合、新しいディレクトリを作成
            switch (Test-Path $savepath) {
                $True { Break }
                $False { New-Item -Path $datapath$foldername -Name "endpoints" -ItemType "directory" }

            }

            ## 最新版のendpointsをダウンロード (noipv6) 
            $url = "$ws/endpoints/Worldwide?noipv6&clientRequestId=$clientRequestId"
            $file = $savepath + $version.latest + ".json"
            $webclient.DownloadFile($url, $file) 

            ## endpointsの変更の履歴をダウンロード (noipv6) 
            $url = "$ws/changes/worldwide/0000000000?noipv6&clientRequestId=$clientRequestId"
            $file = $datapath + $foldername + "history.json"
            $webclient.DownloadFile($url, $file) 

            ## 更新された$version.latestの情報をローカルに保存 
            @($clientRequestId, $version.latest) | Out-File $datapath$foldername$filename

            ## メッセージ：アップデートの完了
            msgCompleted 

        }
        else { 
            ## メッセージ：アップデートの完了
            msgCompleted 
        } 

        Write-Output ""
        Write-Output "動作を選んでください"

        Write-Output "0. プログラムを終了する"
        Write-Output "1. 差分を取得する"

        $answer = $null
        do {
            [String]$answer = Read-Host "入力"

        } while (!($answer -eq 0 -or $answer -eq 1)) 

        Write-Output ""
        switch ($answer) { 
            0 { exit } #switch ($answer) condition 0 end //
            1 { 
                $hPath = $datapath + $foldername + "history.json"
                $history = Get-Content -Path $hPath | ConvertFrom-Json
                $vList = $history | Select-Object -Property "version" -Unique | Sort-Object "version" -Descending

                Write-Output "最新版のエンドポイントと比較するエンドポイントのバージョンを入力してください"
                Write-Output " 0. プログラムを終了する"
                
                $refNum = 0
                for ($i = 1; $i -lt $vList.Count; $i++) {
                    $str = $null
                    $file = $savepath + $vList[$i].version + ".json"

                    if (Test-Path $file) {
                        $refNum += 1

                        switch ($true) {
                            ($refNum -lt 10) { $str = " " + "$refNum" + '. ' + $vList[$i].version }
                            ## ($refNum -ge 10) { $str = "$refNum" + '. ' + $vList[$i].version }
                            ($refNum -eq 10) { $str = " a. さらに表示" }

                        }

                    }
                    Write-Output $str

                }

            } #switch ($answer) condition 1 end //

        } #switch ($answer) end //

        ######TODO#####
        ## 日付を指定して変数に保存
        if ($refNum -lt 10) {
            $iDate = $null
            do {
                $iDate = Read-Host "入力"
    
                if ($iDate -eq 0) { exit }
                else { }
    
            } while (!($iDate -match "^[1-9]+[0-9]*$" -and [Int]$iDate -ge 1 -and [Int]$iDate -le $refNum))

        } else {
            $refNum = 9
            $aSwitch = $False

            $iDate = $null
            do {
                $iDate = Read-Host "入力"
    
                switch ($iDate) {
                    0 { exit }
                    a {
                        if ($aSwitch -eq $False) {
                            $aSwitch = $True
                            Write-Output ""
                            Write-Output "＜バージョン一覧＞"
                            Write-Output " 0. プログラムを終了する"
                            $refNum = 0
                            for ($i = 1; $i -lt $vList.Count; $i++) { 
                                $str = $null 
                                $file = $savepath + $vList[$i].version + ".json"
                    
                                if (Test-Path $file) {
                                    $refNum += 1
                    
                                    switch ($true) {
                                        ($refNum -lt 10) { $str = " " + "$refNum" + '. ' + $vList[$i].version }
                                        ($refNum -ge 10) { $str = "$refNum" + '. ' + $vList[$i].version }
        
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
        $fLatest = $savepath + $version.latest + ".json"
        $eLatest = Get-Content -Path $fLatest | ConvertFrom-Json

        ## 指定した日付のendpointsを呼び出す
        $fIndicated = $savepath + $vList[$iDate].version + ".json"
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
        # $aList.Columns.Add("key(HEX)")
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
            $sIpHex = @($sIp) | ForEach-Object { toHex $_.InputObject } | Sort-Object 
            $sIpDec = @($sIpHex) | ForEach-Object { toDec $_ } 
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

                <#
                        $aList.Rows.Add($no, $impact, "URL", $prop.id, $prop.serviceArea, $prop.serviceAreaDisplayName 
                            , @($sUrl)[$j].InputObject, @($sUrl)[$j].InputObject, $prop.tcpPorts, $prop.udpPorts, $prop.expressRoute 
                            , $prop.category, $prop.required, $prop.notes, $dpChecked) 
                        #>
            
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

                <#
                        $aList.Rows.Add($no, $impact, "IP", $prop.id, $prop.serviceArea, $prop.serviceAreaDisplayName 
                            , @($sIpDec)[$j], @($sIpHex)[$j], $prop.tcpPorts, $prop.udpPorts, $prop.expressRoute 
                            , $prop.category, $prop.required, $prop.notes, $dpChecked) 
                        #>

                $aList.Rows.Add($no, $impact, $dpChecked, "IP", $prop.id, $prop.serviceArea, $prop.serviceAreaDisplayName 
                    , @($sIpDec)[$j], $prop.tcpPorts, $prop.udpPorts, $prop.expressRoute 
                    , $prop.category, $prop.required, $prop.notes) 

            } 

        } 

    } #switch ($direction) condition (1) end //

} #switch ($direction) end //

## 比較した結果を保存するパスを指定 
$savepath = $datapath + $foldername + "Reports" 
if (!(Test-Path $savepath)) { New-Item -Path $datapath$foldername -Name "Reports" -ItemType "directory" } 

$time = Get-Date -Format "yyyyMMdd_HHmmss"
$aList | ConvertTo-Csv -NoTypeInformation -UseCulture | Out-File -FilePath "$savepath\report($time).csv"

#Define locations and delimiter
$csv = "$savepath\report($time).csv" #Location of the source file
$xlsx = "$savepath\report($time).xlsx" #Desired location of output
$delimiter = "," #Specify the delimiter used in the file

# Create a new Excel workbook with one empty sheet
$excel = New-Object -ComObject excel.application 
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)

# Build the QueryTables.Add command and reformat the data
$TxtConnector = ("TEXT;" + $csv)
$Connector = $worksheet.QueryTables.add($TxtConnector, $worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType = 1
$query.TextFileColumnDataTypes = , 1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

# Freeze the top row of the workbook
$excel.Rows.Item("2:2").Select()
$excel.ActiveWindow.FreezePanes = $true

# Set the header background color
$worksheet.Range("A1:N1").interior.colorindex = 20

# Execute & delete the import query
$query.Refresh()
$query.Delete()

# Save & close the Workbook as XLSX.
$Workbook.SaveAs($xlsx, 51)
$excel.Quit()

<#
        $answer = $null 
        $datapath = $null 
        $filename = $null 
        $foldername = $null 
        $savepath = $null 
        $version = $null 
        $ws = $null 
        #>
