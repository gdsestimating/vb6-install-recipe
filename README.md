# VB6 Install Recipe for Windows 10

This was put together from various bits of information I found on the web and a little of my own experimenting. The key was to edit the `.stf` file of vb6 setup files. It is sort of like an early `.msi` format called acme setup. You can change it to effectively select options automatically and skip various parts of the installation that don't work for Windows 10+.

The custom `.stf` file inlcuded here removes VSS and some other tools from the install. I don't totally recall all the options I removed. Once I got it working, I stopped while I was ahead. You can compare it with the original stf file in the vb6 setup files using a diff tool if you really care to see the changes.

## Installing

There are two options for this setup. Docker installs require option 2.


This script assumes you have the VB6 Enterprise .iso and Service Pack 6 exe. Both seem to only be available via https://my.visualstudio.com with Microsoft's Dev Essentials subscription or similar.

### Option 1 Preparation

    Docker does not support iso mounting so if your intent to run it in a docker container you must use option 2 below.

1. Open the Service Pack iso and extract the `Vs6sp6B.exe` file to this project directory on your machine
2. Copy the vb6 cd1 iso to the same directory. Be sure it's named `en_vb6_ent_cd1.iso` or edit the `vb6setup.ps1` script to indicate a different filename.

At the very least, you should run the script from a directory with the following files:

```
  vb6setup.ps1
  vb98ent-custom.stf
  Vs6sp6B.exe
  en_vb6_ent_cd1.iso
```

## Option 2 Preparation

1. Open the Service Pack iso and extract the `Vs6sp6B.exe` file to this project directory on your machine
2. Extract the contents of the vb6 iso to a directory called `vb6` within the project directory. The script will look for the setup exe in that location. If it finds it, it will skip the step of extracting files out of the iso for you. Note: this is different from local installs because we cannot mount .iso files in docker.
3. For docker: Inspect and modify the Dockerfile as needed.

```
  vb6setup.ps1
  vb98ent-custom.stf
  Vs6sp6B.exe
- vb6/
    IsoFilesHere
```

## Run

1. Open a powershell terminal in Administrator mode.
2. Navigate to the directory with the files in it. <-- important. The script looks in the current directory for files.
3. Run the script `> .\vb6setup.ps1` if installing to your current machine OR `docker build` if creating a Docker image.

## Customizing the STF

I think the main thing is to remove 'yes' from some of the high-level items you don't want. I actually removed some of the references to various STF entries instead, but that that was done before I knew to use `/b 2` (Custom install) instead of `/b 1` (Typical install) in the acmsetup.exe args within the script. It worked anyway so I quit while I was ahead. There may be options you want that I excluded in which case you should copy the original from `vb6/Setup/vb98ent.stf` to `vb98ent-custom.stf` and delete the 'yes' for options you don't want. Understand however that some options that I have not specifically identified will cause the install to hang.

More on the stf file format here:
https://fileformats.fandom.com/wiki/Microsoft_ACME_Setup

## Troubleshooting

I've only tested this on Windows 10 Pro. I tried to put comments in the script to document what it does. If you create your own `.stf`, common problems are that the setup hangs toward the end when installing particular optional components. I'm not sure which specific ones are problematic but some SO and forum discussions indicated it is some of the Data Access components.
