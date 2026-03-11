# Run the command and process each line line-by-line
vssadmin list shadowstorage | ForEach-Object {
    $line = $_.Trim()

    # Match and color-code based on the specific storage metric
    if ($line -match "^For volume:") {
        Write-Host $line -ForegroundColor White -BackgroundColor Blue
    }
    elseif ($line -match "^Used Shadow Copy Storage space:") {
        Write-Host $line -ForegroundColor Yellow
    }
    elseif ($line -match "^Allocated Shadow Copy Storage space:") {
        # Highlighting the Allocated Space in Cyan
        Write-Host $line -ForegroundColor Cyan
    }
    elseif ($line -match "^Maximum Shadow Copy Storage space:") {
        Write-Host $line -ForegroundColor Green
    }
    else {
        # Standard white text for volume IDs and general info
        Write-Host $line
    }
}
