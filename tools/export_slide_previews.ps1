$ppt = New-Object -ComObject PowerPoint.Application
$pptPath = 'D:\New folder (5)\VuNguyen\DATN_2252948_NguyenVuNguyen_Defense.pptx'
$pres = $ppt.Presentations.Open($pptPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$outDir = 'D:\New folder (5)\VuNguyen\preview_slides'
if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$pres.Slides.Item(3).Export("$outDir\slide_03_content.png", 'PNG', 1920, 1080)
$pres.Slides.Item(4).Export("$outDir\slide_04_ch1.png", 'PNG', 1920, 1080)
$pres.Slides.Item(14).Export("$outDir\slide_14_ch4_pcb.png", 'PNG', 1920, 1080)
$pres.Slides.Item(21).Export("$outDir\slide_21_ch7_proto.png", 'PNG', 1920, 1080)

$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Get-ChildItem $outDir | Select-Object Name, Length
