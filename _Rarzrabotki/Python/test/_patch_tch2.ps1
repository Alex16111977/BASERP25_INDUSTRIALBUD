$path = "C:\Configuration_downloads\BASERP25\DataProcessors\А_СинхронизироватьДеньги.xml"
$bytes = [System.IO.File]::ReadAllBytes($path)
$hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
$enc = New-Object System.Text.UTF8Encoding($hasBom)
$t = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
if ($t.Contains("<Name>БухСуммаОпределена</Name>")) { Write-Output "SKIP: реквизит уже есть"; exit 0 }

$m = [regex]::Match($t, '(?s)<Attribute uuid="[^"]+">\s*<Properties>\s*<Name>ЕстьРасхождение</Name>.*?</Attribute>')
if (-not $m.Success) { Write-Output "FAIL: шаблон ЕстьРасхождение не найден"; exit 1 }
$tpl = $m.Value

$s = $tpl -replace '<Attribute uuid="[^"]+">', '<Attribute uuid="b7c1e2a4-9f31-4d68-a5c2-1e7f4d90aa33">'
$s = $s.Replace("<Name>ЕстьРасхождение</Name>", "<Name>БухСуммаОпределена</Name>")
$s = [regex]::Replace($s, '<v8:content>[^<]*</v8:content>', '<v8:content>Бух-сума визначена</v8:content>')

$indent = "`r`n" + ("`t" * 5)
$new = $t.Substring(0, $m.Index + $m.Length) + $indent + $s + $t.Substring($m.Index + $m.Length)
[System.IO.File]::WriteAllText($path, $new, $enc)
Write-Output ("OK: добавлен реквизит, размер " + (Get-Item $path).Length)