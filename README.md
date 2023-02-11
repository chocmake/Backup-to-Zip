# Backup to Zip
Generate a shortcut (LNK) configured with your chosen files/directories that when opened will output a timestamped backup zip at your chosen destination.

Intended as a straightforward, quick way for backing up handfuls of files/directories. Originally made for keeping backups of game save files and settings, as a form of automatic versioning for organization and later restoration if needed.
> Since LNK shortcuts have field length limits only so many paths can be included for the backup using this method. See the [Limitations](#limitations) section at the bottom for more details.

## Setup and basic usage

1. Firstly have [7-Zip](https://www.7-zip.org/index.html) installed. As of v0.2 the script will automatically detect it in its default (64-bit) install location.
> If you installed 7-Zip to a alternative location you can add its directory to the Windows PATH, see this [GIF guide](https://user-images.githubusercontent.com/34178938/179670355-82005d39-8277-42cf-a49f-05045e3b8699.gif)
2. [Download](https://github.com/chocmake/Backup-to-Zip/releases/latest/download/Backup.to.Zip.zip) the Backup to Zip script.
3. Open the script and follow the prompts, dragging in the files/directories you want in the backup, along with a destination directory and optional shortcut name, as seen in the GIFs below.
> You can also drag files/directories onto the script's icon and it will auto fill them as input sources.
4. After the prompts have been completed an LNK shortcut will be generated. It's this shortcut that you can then double-click from then on to output a timestamped zip at your chosen destination.

> The shortcut can also be moved elsewhere once created. Additionally the script can be moved if wanted as Windows will update the LNK path automatically to the new location.

## Overview

Showing first setting up the sources and destination with the script, then using the shortcut afterward to auto create the timestamped zips. Here using the default timestamp setting for zip filenames.

![Demo-1](https://user-images.githubusercontent.com/34178938/179670325-24cfa20f-a239-4b8a-b343-c62c27da9365.gif)

Below shows an alternative way to add sources/destination individually into the script window, useful if files/directories are spread across different windows. Here using the date-only setting for zip filenames.

> Also showing if one wants to add comments to zip filenames how the script expects them placed at the end in square brackets.

![Demo-2](https://user-images.githubusercontent.com/34178938/179670339-4cb5fda0-bfac-4c8b-a6c1-7222cd86d984.gif)

## Settings

- Select between zip or 7z archive type for backups.
- Choose to add the time to the filename timestamps (default) or just the date with an increasing counter.
- 12/24 hour time (depending on the former setting).
- Option to additionally preserve the date created and date accessed timestamps for the zip (normally only date modified timestamps are copied).
> Since older versions of 7-Zip lack support for the above feature only enable if you have the 7-Zip v21+ installed.
- Whether to wrap square brackets around the shortcut's filename to raise it above other files when sorted by name.
- Color scheme. Auto/dark/light. Auto (default) will detect the system's theme on W10+ ([screenshot](https://user-images.githubusercontent.com/34178938/179670392-4f23af1f-eaed-4c13-bbf3-7bc45f90020f.png)).

## Other features

- The original source(s)/destination paths from a LNK shortcut can be extracted to a text file by dragging the shortcut onto the script. The text file will be created beside the shortcut:

![Demo-3](https://user-images.githubusercontent.com/34178938/179670347-6faec160-1bdd-4bcd-b970-afeb6f719e22.gif)

- If 7-Zip reports that some file(s) couldn't be included in the zip (due to eg: an application in use locking them, or they were renamed or moved) the script will list which file(s) are missing and add `[m]` to the filename of the zip.

- If the entered Destination path doesn't exist (eg: if it's a file instead of a directory or the path doesn't exist, such as when typing/pasting one in) the script will re-prompt for a valid directory.

## Limitations

- LNK files have a character limit of ~1023 for its arguments field (determined after more tests). This means there are only so many paths that can be used for input sources and if wanting a larger number of files in the backup it's best to add the directory containing them (counts as single path) rather than a mass of individual files themselves.
- The native Windows method of generating LNK shortcuts, WScript.Shell, has issues with very specific Unicode characters in paths, specifically characters that *appear like* invalid filename characters (`?`, `/`, `:`) but which are in fact valid (such as U+FF1F, U+2215, U+003A). Since v0.2 of the script it now has a workaround for this Windows limitation so any paths containing those specific Unicode characters are correctly handled for the backup.
- It doesn't yet check for whether a Destination directory has write permissions for the current user and will fail to initially create the LNK if it can't write to it.
- Only works for paths with drive letters currently (ie: won't work on unmapped network paths).
