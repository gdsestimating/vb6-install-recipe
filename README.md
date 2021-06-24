# VB6 Install Recipe for Windows 10

This was put together from various bits of information I found on the web and a little of my own experimenting. The key is to edit the `.stf` file of vb6 setup files. It is sort of like an early `.msi` format. You can change it to effectively select options automatically and skip various parts of the installation that don't work for Windows 10+.

The custom `.stf` file inlcuded here removes VSS and some other tools from the install. I don't totally recall all the options I removed. Once I got it working, I stopped while I was ahead. You can compare it with the original stf file in the vb6 setup files using a diff tool if you really care to see the changes.

## Prerequisites

This script assumes you have the VB6 Enterprise .iso and Service Pack 6 exe. Both seem to only be available via https://my.visualstudio.com with Microsoft's Dev Essentials subscription.

1. Open the Service Pack iso and extract the `Vs6sp6B.exe` file to this project directory on your machine
2. Copy the vb6 cd1 iso to the same directory. Be sure it's named `en_vb6_ent_cd1.iso` or edit the `vb6setup.ps1` script

**OR**

2.  Extract the contents of the vb6 iso to a directory called `vb6` within the project directory. The script will look for the setup exe in that location. If it finds it, it will skip the step of extracting files out of the iso for you.

At the very least, you should have a directory structure like:

```
- vb6setup.ps1
- vb98ent-custom.stf
- Vs6sp6B.exe
- en_vb6_ent_cd1.iso
```

## Run

1. Open a powershell terminal in Administrator mode.
2. Navigate to the directory with the files in it.
3. Run the script

## Troubleshooting

I've only tested this on Windows 10 Pro. I tried to put comments in the script to document what it does. If you create your own `.stf`, common problems are that the setup hangs toward the end when installing particular components. I'm not sure which specific ones are problematic. Some info on that can be found in stack overflow and some old blogs.