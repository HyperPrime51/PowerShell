# variables
$source = "F:"
$target = "\\rds-03\c$\users\clee"
$target_drive = "G:"

# desktop
Copy-item -Force -Recurse -Verbose -Path "$($source)\desktop" -Destination "$($target)"

# documents
Copy-item -Force -Recurse -Verbose -Path "$($source)\documents" -Destination "$($target)"

# downloads
Copy-item -Force -Recurse -Verbose -Path "$($source)\downloads" -Destination "$($target)"

# favorites
Copy-item -Force -Recurse -Verbose -Path "$($source)\favorites" -Destination "$($target)"

# chrome favorites
Copy-item -Force -Verbose -Path "$($source)\AppData\Local\Google\Chrome\User Data\Default\Bookmarks" -Destination "$($target)\AppData\Local\Google\Chrome\User Data\Bookmarks"
Copy-item -Force -Verbose -Path "$($source)\AppData\Local\Google\Chrome\User Data\Default\Bookmarks.bak" -Destination "$($target)\AppData\Local\Google\Chrome\User Data\Bookmarks.bak"

# pinned folders
Copy-item -Force -Recurse -Verbose -Path "$($source)\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations" -Destination "$($target)\AppData\Roaming\Microsoft\Windows\Recent"

# pinned app in taskbar
# export
Import-RegistryHive -File "$($source)\NTUSER.DAT" -Key 'HKU\TEMP_SOURCE_HIVE' -Name TempSourceHive
REG EXPORT HKEY_USERS\TempSourceHive\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband "$($target)\tb-pinned-items.reg"
# replace HKEY_USERS\TempSourceHive with HKEY_CURRENT_USER
Copy-item -Force -Recurse -Verbose -Path "$($source)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" -Destination "$($target)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned"
Remove-RegistryHive -Name TempSourceHive

# import
# needs to run under user profile
REGEDIT /S "%userprofile%\tb-pinned-items.reg"

# new teams
# path to startup C:\Users\clee\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup