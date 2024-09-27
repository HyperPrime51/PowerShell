# variables
$source = "F:"
$target = "\\rds-02\c$\users\clee"
$target_drive = "G:"

# desktop
Copy-item -Force -Recurse -Verbose -Path "$($source)\desktop" -Destination "$($target)\desktop"

# documents
Copy-item -Force -Recurse -Verbose -Path "$($source)\documents" -Destination "$($target)\documents"

# downloads
Copy-item -Force -Recurse -Verbose -Path "$($source)\downloads" -Destination "$($target)\downloads"

# favorites
Copy-item -Force -Recurse -Verbose -Path "$($source)\favorites" -Destination "$($target)\favorites"

# chrome favorites
Copy-item -Force -Verbose -Path "$($source)\AppData\Local\Google\Chrome\User Data\Default\Bookmarks" -Destination "$($target)\AppData\Local\Google\Chrome\User Data\Bookmarks"
Copy-item -Force -Verbose -Path "$($source)\AppData\Local\Google\Chrome\User Data\Default\Bookmarks.bak" -Destination "$($target)\AppData\Local\Google\Chrome\User Data\Bookmarks.bak"

# pinned folders
Copy-item -Force -Recurse -Verbose -Path "$($source)\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations" -Destination "$($target)\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations"

# pinned app in taskbar
# export
Import-RegistryHive -File "$($source)\NTUSER.DAT" -Key 'HKU\TEMP_SOURCE_HIVE' -Name TempSourceHive
REG EXPORT HKEY_USER\TEMP_HIVE\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband "$($target)\tb-pinned-items.reg"
Copy-item -Force -Recurse -Verbose -Path "$($source)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" -Destination "$($target)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
Remove-RegistryHive -Name TempSourceHive

# import
# needs to run under user profile
REGEDIT /S "%userprofile%\tb-pinned-items.reg"