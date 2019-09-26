# RIP Ripper

The RIP Ripper is a modulized Visual Basic Script for automating the separation and formatting of most of the products received from the Enterprise Output Manager used by Personnel Systems Managers in the United States Air Force.

  - Open-source script (always has been, as is the nature of scripts)
  - No installation required
  - Extremely customizable configuration
  - Expandable and flexible to accomodate new products
  - Used across multiple MAJCOMS
  - Creator coined by AETC/A1 for creating it
  - Can save every Military Personnel Flight hours of work DAILY
  - Can be automated using Windows Scheduled Tasks
  - Multiple methods to utilize the script available by default

Why Visual Basic Script (aka VBS)? Because the Air Force network has such a high security requirement that I'm shocked even this works. As it is, Windows comes with a built-in script processor for VBS that it relied on for many important functions. This allows us to take advantage of that same script processor with our own scripts. The day they block VBS scripts is the day Windows will cease to function on the AFNET.

### Unofficial Official Support

I have attempted to get AFPC to take official ownership of these scripts and update/provide/integrate them on official channels, but it was turned down. They did pay me some money through the Airmen Powered by Innovation (API) program, however, so I won't complain.

I keep busy, but I've yet to turn away someone asking for help. If you're having issues with this script or with EOM, swing by the [issues section](https://gitlab.com/usaf-psm/rip-ripper/issues) and let me know what's up. No problem is too big or too small! If you would just like to request a single product be added to the config and can't figure it out yourself, I don't mind. If you would like help customizing your own setup to automate your entire product line, I may be able to offer some help with that too!

### How to Deploy

The "Hook Method" is the most common usage of the ripper. The steps below will lead you to creating a "hook" file. You can then move that file to anywhere that you have EOM (.BKP) files stored and execute it to process all the EOM files sharing the same folder as the hook file itself. I call it the "hook" method and the "hook" file because it literally is a small script that reaches back to the original location of the main script, hooks into the main script, and then points it back to where the small script started.

Requirements:

 - Windows 7 or above
 - Microsoft Word 2010 or above
 - Enterprise Output Manager (EOM) configured to output backup (.BKP) files.

Setup Steps (Hook Method): 

 1. Download [all the files](https://gitlab.com/usaf-psm/rip-ripper/-/archive/master/rip-ripper-master.zip) in this repository
 2. Place these files in a location that is accessible by everyone who intends to use the ripper, but preferably not where you keep your EOM files
 3. Execute the "RIP Ripper Make Hook.vbs" script.
 4. Move the newly created "RIP Ripper Hook.bat" to anywhere you may have .BKP files that you want to be processed by the ripper.

### Post Script
Another method that's currently used at Sheppard AFB is the "Scheduled Method". This requires some knowledge of Windows' Scheduled Tasks, such as how to create and modify them, in order to get it working correctly. I don't provide instructions for this method as it is extremely complex and not recommended for your average user.

If you're feeling daring, however, you would simply make a call to the "RIP Ripper Core.vbs" file in your scheduled task and pass it a folder location as a single argument. Like so: `"C:\ripper\RIP Ripper Core.vbs" C:\EOM\`

This executes the ripper at `C:\ripper\` and then tells it where to look for the BKP files which, in this example, are at `C:\EOM\`. This is exactly what the hook does, you would just be doing it yourself.