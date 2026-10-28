# 从B站公开接口挑选适合 E2E 测试的视频（整页PPT型演讲/课程）
$UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
$H = @{ 'User-Agent' = $UA; 'Referer' = 'https://www.bilibili.com/' }

Write-Host "===== 1) 搜索页抓取：王道考研/操作系统（纯PPT课程类）====="
try {
    $html = (Invoke-WebRequest -Uri 'https://search.bilibili.com/all?keyword=%E7%8E%8B%E9%81%93%E8%80%83%E7%A0%94%20%E6%93%8D%E4%BD%9C%E7%B3%BB%E7%BB%9F' -Headers $H -TimeoutSec 20 -UseBasicParsing).Content
    $bvs = [regex]::Matches($html, 'BV[0-9A-Za-z]{10}') | ForEach-Object { $_.Value } | Select-Object -Unique -First 10
    Write-Host ("搜索页命中 {0} 个BV" -f @($bvs).Count)
    $bvs | ForEach-Object { Write-Host "  BV候选: $_" }
    # 提取标题（用于判断）
    $titles = [regex]::Matches($html, '"title":"([^"]{5,80}?)"') | ForEach-Object { $_.Groups[1].Value } | Select-Object -First 15
    $titles | ForEach-Object { Write-Host "  TITLE: $_" }
} catch {
    Write-Host "搜索页抓取失败: $($_.Exception.Message)"
}

Write-Host "===== 2) 科技区排行榜（找 发布会/演讲/公开课）====="
try {
    $r = Invoke-RestMethod -Uri 'https://api.bilibili.com/x/web-interface/ranking/v2?rid=36&type=all' -Headers $H -TimeoutSec 20
    $r.data.list | Where-Object { $_.title -match '发布会|演讲|公开课|讲座|大会|keynote|Keynote' } | Select-Object -First 10 |
      ForEach-Object { "{0} | {1}分{2}秒 | {3} | UP:{4}" -f $_.bvid, [math]::Floor($_.duration/60), ($_.duration%60), (($_.title -replace '<[^>]+>','')), $_.owner.name }
} catch {
    Write-Host "排行榜失败: $($_.Exception.Message)"
}
