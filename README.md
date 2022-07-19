# Backup to Zip
Generate a shortcut (LNK) that when opened creates an auto timestamped ZIP using 7-Zip of the specified source files at a destination. Made originally as a straightforward, quick way for versioning game save file directories and config files.

## Setup and basic usage

1. Firstly have [7-Zip](https://www.7-zip.org/index.html) installed, then add its directory to the Windows PATH environment variable for your account so the script can find it (follow [this GIF guide](https://user-images.githubusercontent.com/34178938/179670355-82005d39-8277-42cf-a49f-05045e3b8699.gif)).
2. Download the Backup to Zip script either above or in zipped form [here](https://github.com/chocmake/Backup-to-Zip/releases/download/0.1/0.1.2022-07-18.zip).
3. Then set up the Source/Destination and optional shortcut name using the script as seen in the GIFs below. Once the shortcut is made you just double-click it from then on to output a timestamped zip.

> The shortcut can also be moved elsewhere once created, just remember to keep the script in the same location so the shortcut can find it.

## Overview

Showing first setting up the sources and destination with the script, then using the shortcut afterward to auto create the timestamped zips. Here using the default timestamp setting for zip filenames.

![Demo-1](https://user-images.githubusercontent.com/34178938/179670325-24cfa20f-a239-4b8a-b343-c62c27da9365.gif)

Alternative way to add sources/destination individually into the script window, useful if files/directories are spread across different windows. Here using the date-only setting for zip filenames. Also showing if one wants to add comments to zip filenames how the script expects them placed at the end in square brackets.

![Demo-2](https://user-images.githubusercontent.com/34178938/179670339-4cb5fda0-bfac-4c8b-a6c1-7222cd86d984.gif)

## Additional features

- Settings:
  - Choose to add the time to the filename timestamps (default) or just the date with an increasing counter.
  - 12/24 hour time (depending on the former setting).
  - Whether to wrap square brackets around the shortcut's filename to raise it above other files when sorted by name.
  - Color scheme. Auto/dark/light. Auto (default) will detect the system's theme on W10+:

![Auto-theme-detection](https://user-images.githubusercontent.com/34178938/179670392-4f23af1f-eaed-4c13-bbf3-7bc45f90020f.png)

- To extract the original source(s)/destination paths from a LNK shortcut it can be dragged onto the script and a text file will be output beside it containing the paths for reference. GIF:

![Demo-3](https://user-images.githubusercontent.com/34178938/179670347-6faec160-1bdd-4bcd-b970-afeb6f719e22.gif)

## Limitations

- LNK files have limits to their field length of ~32k afaict (used by the script to store the sources/destination/custom name) and CMD itself has its own max buffer size. Something to keep in mind as the use case for the script was originally for mostly directory paths of game saves (which are typically just the parent directory paths and some single files, rather than hundreds/thousands of individual source paths entered which would need to be stored in the LNK's field).
- The common, widely used method used to generate the LNK shortcut (WScript.Shell) fails to work with some very specific Unicode characters (in my testing U+FF1F and U+2215), while being fine with everything else I've tried.
- It doesn't yet check for whether a Destination directory has write permissions for the current user and will fail to initially create the LNK if it can't write to it.
- Only works for paths with drive letters currently (ie: won't work on unmapped network paths).
