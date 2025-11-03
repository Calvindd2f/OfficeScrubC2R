$isInstalled? = gcim win32_product | where-object { $_.Name -like "*Office*" }
if ($isInstalled?) { write-host "Office is installed" } else { write-host "Office is not installed"}

