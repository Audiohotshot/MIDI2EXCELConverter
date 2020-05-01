# MIDI2EXCELConverter
v6.0 

Convert MIDI files to XLS sheets with this VBA macro for Excel 

Author: www.Audiohotshot.nl

*PLEASE DO NOT CHANGE THIS FILE, IT'S GOING TO BE OVERWRITEN ON EACH UPGRADE*

# Purpose:
Make MIDI files readable in EXCEL for researchpurposes. Shows timing, delta timing and velocity values; these values 
can then be easily used in other calculations for researchpurposes.

# Installation:
Copy all files to your computer 

# Launch plugin:
Open the .xlsm file.
It contains macros; so allow to use macros in excel.

There's just one big start button. Press it and a finder/explorer will pop up for loading the .midi file.
This could be done in a few seconds up to a minute, depending on the size of the file.
While processing the screen could be empty.
Let it run, it will show everything at once.

# Features:
The macro captures:
META related:
-   Timing
-   Division
-   Tracknames
-   Tracklength
-   TEXT information
-   File format

MIDI related:
-   Note on/off
-   Sysex
-   Aftertouch (poly and channel)
-   Pitchbend
-   Control Changes
-   Programchanges
-   Text

This was developed on a Mac running Excel v15 and Sublime Text 3

# Known issues:
Not all MIDI features or META features are implemented, just the most common.

# Disclaimer:
The software is made available "AS IS". It seems quite stable, but it is in
an early stage of development.  Hence there should be plenty of bugs not yet
spotted. This is for educational purpose only.
