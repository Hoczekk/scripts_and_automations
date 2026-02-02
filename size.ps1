$size = 0 
foreach ($file in (get-childitem <folder -file)> {$size += $file.length} 
$size
