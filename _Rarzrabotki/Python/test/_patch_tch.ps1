$path = "C:\Configuration_downloads\BASERP25\DataProcessors\А_СинхронизироватьДеньги.xml"
$bytes = [System.IO.File]::ReadAllBytes($path)
$hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
$enc = New-Object System.Text.UTF8Encoding($hasBom)
$t = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
Write-Output ("BOM: " + $hasBom + "; CRLF: " + $t.Contains("`r`n"))

if ($t.Contains("<Name>ПоясненоДокументами</Name>")) { Write-Output "SKIP: реквизиты уже есть"; exit 0 }

$m = [regex]::Match($t, '(?s)<Attribute uuid="[^"]+">\s*<Properties>\s*<Name>Разница</Name>.*?</Attribute>')
if (-not $m.Success) { Write-Output "FAIL: шаблон Разница не найден"; exit 1 }
$tpl = $m.Value

function New-Attr($tpl, $uuid, $name, $ru, $uk) {
    $s = $tpl -replace '<Attribute uuid="[^"]+">', ('<Attribute uuid="' + $uuid + '">')
    $s = $s.Replace("<Name>Разница</Name>", "<Name>$name</Name>")
    $s = $s.Replace("<v8:content>Разница</v8:content>", "<v8:content>$ru</v8:content>")
    $s = $s.Replace("<v8:content>Різниця</v8:content>", "<v8:content>$uk</v8:content>")
    return $s
}

$a1 = New-Attr $tpl "b7c1e2a4-9f31-4d68-a5c2-1e7f4d90aa11" "ПоясненоДокументами" "Пояснено документами" "Пояснено документами"
$a2 = New-Attr $tpl "b7c1e2a4-9f31-4d68-a5c2-1e7f4d90aa22" "НеПояснено" "Не пояснено" "Не пояснено"

$indent = "`r`n" + ("`t" * 5)
$new = $t.Substring(0, $m.Index + $m.Length) + $indent + $a1 + $indent + $a2 + $t.Substring($m.Index + $m.Length)
[System.IO.File]::WriteAllText($path, $new, $enc)
Write-Output ("OK: добавлено 2 реквизита, размер " + (Get-Item $path).Length)