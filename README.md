
<p align="center">
<img id="charmsbarPort" src="resource/darklogo.png"/>
</p>

<blockquote>
"Your most unhappy customers are your greatest source of learning." -Bill Gates
</blockquote>

## About
<b>Charms Bar Port</b> is the brand new solution for bringing back the Windows 8.x Charms bar to Windows 10 and Windows 11, using real files from Windows 8.x to meet your cravings and enhance your desktop.

Forked and completely edited from <a href="https://github.com/Jerhynh/CharmsBarRevived">CharmsBarRevived</a>, <b>Charms Bar Port</b> will assist on helping you transition to Windows 10 and 11 without having to keep on the obsolete system forever.

## Why was this created?
As you may know, ValiNet Charms is no longer available to download and an Archive.org repository containing a newer version appears to be in Russian and is infected, so I created this project primarily to bring my needs of a Charms Bar back.

## How does it work?
On touch screens, swipe from the right edge towards to bring up the Charms bar. If you're a mouse user, swipe to the top right corner and drag your cursor down to open the Charms bar. You can also use the keyboard shortcut Windows key + C, just like it was on Windows 8.x.

## Features
* Powered by Visual Studio 2022
* Based on Windows 8.1 Update 3
* Includes accent colors
* Can read your network/battery status
* Supports Windows 8.x-era registry hacks (DisableTRCorner and DisableBRCorner)

## Screenshots
<img src="resource/preview.png"/>

## Download
Downloads are coming soon in the near future

## Q&As
Q: When will this be released?<br />
A: I personally don't know as I'm always busy with other things.
<br />
<br />
Q: Why is there no touch screen support?<br />
A: I do not have access to a Windows tablet to add it.  If anyone wants to have their go at adding this in, then ask me, and I'll send a Google Drive link for you to improve upon.
<br />
<br />
Q: How can I disable the Charms Bar hot corners without closing the program?<br />
A: This requires fiddling with the registry. I am not responsible if you mess up your system.

1. Press the “WIN+R” key combination to launch the Run dialog box, then type regedit and press enter. It’ll open the Registry Editor, and go to following key:
<blockquote>
HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ImmersiveShell\
</blockquote>
<br />
2. Under the ImmersiveShell key, create a new key called EdgeUI.
<br />
3. Now select the newly created key “EdgeUI” and in the right-side pane, create two new DWORDs named DisableTRCorner and DisableBRCorner and set their values to 1.
<br />
4. That’s it. It’ll immediately disable the Charms Bar hot corners. You do not need to log off or restart the system.
