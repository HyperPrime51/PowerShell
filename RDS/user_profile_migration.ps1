# variables
$source = "F:\"
$target = "\\rds-02\c$\users\clee"

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



# pinned app in taskbar, needs to be run on user profile
# export
REG EXPORT HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband "%userprofile%\tb-pinned-items.reg"
xcopy "%AppData%\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" "%userprofile%\Pinned Items Backup\pinnedshortcuts" /E /C /H /R /K /Y

# import
REGEDIT /S "%userprofile%\tb-pinned-items.reg"
xcopy "%userprofile%\Pinned Items Backup\pinnedshortcuts" "%AppData%\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" /E /C /H /R /K /Y