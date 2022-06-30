# PeoplePicker Instructions

## ChangeLog

* Modified code to handle querying any entity type. Custom coded it to handle accounts, users, customers, and teams as well.
* Modified it to handle disabled/readonly. Before if you opened an deactivated record - it would still allow editing
* Modified it to pass third parameter back as part of the object - entityName. Is useful for parsing data for reports/etc.
* Added buildscript to automate the build process. Just need to open PowerShell as admin, navigate to buildscript.ps1 and then type that in and press enter. It will clean up previous files and build/run all necessary steps once you have the initial setup done. 

## Pre-Req's

Pre-req
Artilce reference - https://debajmecrm.com/2019/04/26/part-2-setting-up-development-environment-and-custom-component-project/ (please note, some steps are out of date, see below for updated CLI references)

1. Install NPM – https://nodejs.org/en/
2. Install Power Apps CLI - http://download.microsoft.com/download/D/B/E/DBE69906-B4DA-471C-8960-092AB955C681/powerapps-cli-0.1.51.msi.

## Setup Project

1. Create new folder - C:\Controls\OfficeUIFabricReactPeoplePicker-master
2. Open Developer Command Prompt. Switch to folder in step 1.
3. pac pcf init --namespace Ramakrishnan --name OfficeUIFabricReactPeoplePicker --template field
4. run npm install
5. Install typescript - npm install --save-dev typescript

Open VS Command Prompt

1. Navigate to where index.ts is (C:\Controls\OfficeUIFabricReactPeoplePicker-master\OfficeUIFabricReactPeoplePicker)
CMD:npm run build
optional -npm start

2. Change to folder where you want to create zip (i.e. OfficeUIFabricReactPeoplePickerPCF or Deployment)

CMD:pac solution init --publisher-name <enter your publisher name> --publisher-rrefix <enter your publisher name>

pac solution init --publisher-name Ramakrishnan --publisher-prefix rrpcf

3. Add reference
pac solution add-reference --path <path or relative path of your PowerApps component framework project on disk

pac solution add-reference --path C:\Source\repos\controls\OfficeUIFabricReactPeoplePicker-master

4. Build
MSBUILD /t:restore
MSBUILD

5. Zip will be in the folder in step 2, bin folder.

*Note: must delete files from folder in Step 2 whenever you do another build, otherwise you'll get error.

# OfficeUIFabricReactPeoplePicker

A PCF component for Multi select people picker. 

This solution gives the ability to select multiple members from the people picker. It works on all entities/tables of common data services.

Since multiselect peoplepicker is not there yet, this can be used as an alternate. It is build on top of Multiline text field.

# Dependencies
office-ui-fabric-react : https://github.com/OfficeDev/office-ui-fabric-react

# Reference 

Build using this example : https://developer.microsoft.com/en-us/fabric#/controls/web/peoplepicker

## Usage

### Step 1 - Import the solution

  Option 1 - Import the zip file directly into CDS. Managed or Unmanaged Solution.

  Option 2 
  - git clone the repo
  - npm install
  - npm run build
  
 ### Step 2 - Add the component
 Create a Multiline text field and add the component.

## Glimpse of the sample 

![](assets/Peoplepicker.gif)

## For packaging & deploying the solution, refer the below link

 https://docs.microsoft.com/en-us/powerapps/developer/component-framework/import-custom-controls 

## Solution

Solution|Author(s)
--------|---------
OfficeUIFabricReactPeoplePicker|Ramakrishnan Raman

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
