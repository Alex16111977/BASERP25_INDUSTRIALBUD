$path = "C:\Configuration_downloads\BASERP25\DataProcessors\А_СинхронизироватьДеньги\Forms\Форма\Ext\Form.xml"
$bytes = [System.IO.File]::ReadAllBytes($path)
$hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
$enc = New-Object System.Text.UTF8Encoding($hasBom)
$t = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
if ($t.Contains('name="ТРПояснено"')) { Write-Output "SKIP: колонки уже есть"; exit 0 }

$anchor = '<ExtendedTooltip name="ТРРазницаExtendedTooltip" id="82"/>' + "`r`n" + ("`t" * 4) + '</InputField>'
$pos = $t.IndexOf($anchor)
if ($pos -lt 0) { Write-Output "FAIL: якорь ТРРазница не найден"; exit 1 }
$pos = $pos + $anchor.Length

$T4 = "`t" * 4
$T5 = "`t" * 5
$T6 = "`t" * 6
$T7 = "`t" * 7

function New-Col($nm, $path1, $title, $w, $idBase) {
    $s  = "`r`n" + $T4 + '<InputField name="' + $nm + '" id="' + $idBase + '">'
    $s += "`r`n" + $T5 + '<DataPath>Объект.ТаблицаРасхождений.' + $path1 + '</DataPath>'
    $s += "`r`n" + $T5 + '<ReadOnly>true</ReadOnly>'
    $s += "`r`n" + $T5 + '<Title>'
    $s += "`r`n" + $T6 + '<v8:item>'
    $s += "`r`n" + $T7 + '<v8:lang>ru</v8:lang>'
    $s += "`r`n" + $T7 + '<v8:content>' + $title + '</v8:content>'
    $s += "`r`n" + $T6 + '</v8:item>'
    $s += "`r`n" + $T5 + '</Title>'
    $s += "`r`n" + $T5 + '<FooterDataPath>Объект.ТаблицаРасхождений.Total' + $path1 + '</FooterDataPath>'
    $s += "`r`n" + $T5 + '<Width>' + $w + '</Width>'
    $s += "`r`n" + $T5 + '<MarkNegatives>true</MarkNegatives>'
    $s += "`r`n" + $T5 + '<ContextMenu name="' + $nm + 'КонтекстноеМеню" id="' + ($idBase + 1) + '"/>'
    $s += "`r`n" + $T5 + '<ExtendedTooltip name="' + $nm + 'ExtendedTooltip" id="' + ($idBase + 2) + '"/>'
    $s += "`r`n" + $T4 + '</InputField>'
    return $s
}

$c1 = New-Col "ТРПояснено" "ПоясненоДокументами" "Пояснено док." 14 400
$c2 = New-Col "ТРНеПояснено" "НеПояснено" "НЕ пояснено" 14 403

$new = $t.Substring(0, $pos) + $c1 + $c2 + $t.Substring($pos)
[System.IO.File]::WriteAllText($path, $new, $enc)
Write-Output ("OK: добавлено 2 колонки, размер " + (Get-Item $path).Length)