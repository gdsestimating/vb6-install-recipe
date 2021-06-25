# VB6 Install Recipe for Windows 10

This was put together from various bits of information I found on the web and a little of my own experimenting. The key was to edit the `.stf` file of vb6 setup files. It is sort of like an early `.msi` format called acme setup. You can change it to effectively select options automatically and skip various parts of the installation that don't work for Windows 10+.

The custom `.stf` file inlcuded here removes VSS and some other tools from the install. I don't totally recall all the options I removed. Once I got it working, I stopped while I was ahead. You can compare it with the original stf file in the vb6 setup files using a diff tool if you really care to see the changes.

## Prerequisites

This script assumes you have the VB6 Enterprise .iso and Service Pack 6 exe. Both seem to only be available via https://my.visualstudio.com with Microsoft's Dev Essentials subscription.

1. Open the Service Pack iso and extract the `Vs6sp6B.exe` file to this project directory on your machine
2. Copy the vb6 cd1 iso to the same directory. Be sure it's named `en_vb6_ent_cd1.iso` or edit the `vb6setup.ps1` script

**OR**

2.  Extract the contents of the vb6 iso to a directory called `vb6` within the project directory. The script will look for the setup exe in that location. If it finds it, it will skip the step of extracting files out of the iso for you.

At the very least, you should run the script from a directory with the following files:

```
  vb6setup.ps1
  vb98ent-custom.stf
  Vs6sp6B.exe
  en_vb6_ent_cd1.iso
```

Docker does not support iso mounting so if you intent to run it in a docker container you must extract the Visual Basic 6 iso files to a folder called `vb6` like:

```
  vb6setup.ps1
  vb98ent-custom.stf
  Vs6sp6B.exe
- vb6/
    IsoFilesHere
```

## Run

1. Open a powershell terminal in Administrator mode.
2. Navigate to the directory with the files in it.
3. Run the script

## Customizing the STF

I think the main thing is to remove 'yes' from some of the high-level items you don't want. I actually removed some of the references to various STF entries in the current custom file. That was done before I knew to use `/b 2` instead of `/b 1` in the acmsetup.exe args.

https://fileformats.fandom.com/wiki/Microsoft_ACME_Setup

## Troubleshooting

I've only tested this on Windows 10 Pro. I tried to put comments in the script to document what it does. If you create your own `.stf`, common problems are that the setup hangs toward the end when installing particular optional components. I'm not sure which specific ones are problematic but some SO and forum discussions indicated it is some of the Data Access components.
