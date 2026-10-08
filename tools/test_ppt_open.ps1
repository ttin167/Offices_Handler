param([string]$Path = "D:\New folder (5)\VuNguyen\DATN_2252948_NguyenVuNguyen_Defense.pptx")
$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($Path, 0, 0, 0)
    Write-Host "SUCCESS: Slide count is $($pres.Slides.Count)"
    $pres.Close()
} catch {
    Write-Host "CAUGHT EXCEPTION: $($_.Exception.ToString())"
} finally {
    $ppt.Quit()
}
