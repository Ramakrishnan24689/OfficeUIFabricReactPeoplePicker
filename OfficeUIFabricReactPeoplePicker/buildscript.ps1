Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Step 1 - run build
cd C:\Controls\OfficeUIFabricReactPeoplePicker-master\OfficeUIFabricReactPeoplePicker
npm run build

# Step 2 - remove build directory and recreate
Remove-Item -Force -Recurse -Path "C:\Controls\OfficeUIFabricReactPeoplePicker-master\OfficeUIFabricReactPeoplePicker\OfficeUIFabricReactPeoplePickerPCF\*"

# Step 3 - change directory and run processes. 
cd OfficeUIFabricReactPeoplePickerPCF
pac solution init --publisher-name Ramakrishnan --publisher-prefix rrpcf
pac solution add-reference --path C:\Controls\OfficeUIFabricReactPeoplePicker-master

# Step 4 - run builds
MSBUILD /t:restore
MSBUILD

#Step 5 - change back to org folder
cd C:\Controls\OfficeUIFabricReactPeoplePicker-master\OfficeUIFabricReactPeoplePicker
PAUSE
