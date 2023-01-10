This utility is a functional reproduction of Rocketdock. The design is limited 
to emulating what Rocketdock already provides. This will make the utility 
familiar to Rocketdock users.

SteamyDock is Alpha-grade software, under development, not yet ready to use on a 
production system - use at your own risk. Things SteamyDock cant yet do... or 
has problems with:

o Animations are generally not as good as those in RocketDock
o Animations are missing on icon deletion, I aim to implement this in the future.
o Animations are missing on dock closure, I aim to implement this in the future.
o Similar animations to Rocketdocks are missing on drag and drop into the dock, 
    I aim to implement the above in the future.
o SteamyDock is incompatible with Rocketdocks themed backgrounds but instead has 
    its own themeing implementation.
o Cannot use docklets. Docklets derive from an undocumented Objectdock standard 
    and are now obsolete and so Steamydock is unlikely to ever support them.
o The dock itself cannot yet extract embedded icons from EXEs and display them 
    with a transparent background - this is tricky as VB6 has no native PNG 
    support. I aim to implement this in the future when I know how...
o Right/left dock positions will not be supported unless someone shouts loud.
o Only one dock animation type currently implemented.
o Does not support multiple languages in the translation of menu options. 

SteamyDock improves on Rocketdock in various ways:

o For compatibility, the dock can read Rocketdocks obsolete registry/settings 
    locations but now stores its data in the future-proof location of AppData.
o It animates and displays the dock efficiently using less cpu than Rocketdock.
o It is open source which means that anyone can build it, even you.
o It contains no third-party libraries nor any code that is not freely FOSS.
o It is written using a friendly language accessible to beginners and experts.
o The docks associated utilities offer more options for configuration.
o This version is currently supported - by me, you and anyone who cares 
    to make changes or improvements.
o It will soon have a VB.NET/TwinBasic/RadBasic counterpart for 64bit 
    future-proofing.
o It has been tested and runs on Windows XP, Win7 to 10 32bit (and 64 bit in .NET 
    form) and also on 32bit ReactOS.
o It works on the latest version of Windows without compromise in functionality.
o Steamydock allows for increased icon maximum sizes of up to 256x256 pixels, 
    previously Rocketdock was limited to 128,
o Steamydock is fully documented.
o Runs applications as administrator if required.
o Adding icons via the right click menu offers many more icon options.
o Additional menu option for deleting any running application instance.
o Additional hiding options, a new fade, instant and continuous hide.
o Can run new instances of an application at will
o Can maximise and minimise windows, even to/from the systray.
o Can alter the z-order of Windows, sending them to front and back.
o There are tooltips for all controls in the two utilities.
o There is now a readily available help facility, you are reading it...
o The dock itself is fully functional while the configuration utility is 
    operating, Rocketdock caused the dock to be inoperative whilst running.

NOTE: SteamyDock is Alpha-grade software, under development, use at your own risk.

Regardless of what SteamyDock can or cannot do, SteamyDock is FOSS, ie, free and
open source. This means you can download and modify this dock yourself. As long
as there are VB6 programmers willing to fix or change the code then unlike 
Rocketdock, this dock never needs to be made obsolete. If there is something 
broken or missing then you can fix it yourself. So far, Ive managed to make a 
functioning dock and associated programs but be aware it is not yet fully 
complete and not tested in all scenarios. If it worries you to run unfinished 
software on your desktop computer then remove it now please but perhaps come 
back another day when it is finally finished.

This is the fifth VB6 project that I have undertaken and completedso forgive 
the errors in coding styles and methods. Entirely self-taught and a mere 
hobbyist in VB6.

The reason I created it was to teach myself VB6, to get back into the groove.
Back in the 90s I was programming in QB45 and VB DOS and then VB6 but left VB6 
and abandoned my main project when VB6 was deprecated. My skills were paltry 
then and were picked up from the days of Sinclair Zx80s. My aim now is to 
resurrect such skills that I had and improve upon them. A secondary aim is to 
teach myself how to code in technologies that I have encountered. 

When this 
project is complete my next aim is to migrate it to VB.NET through the versions 
to find out what problems are typically encountered in a project such as this. 
 
Starting with VB6, it was a big surprise to me to find such inadequate native 
image type handling, VB6 being unable to handle the various image types without 
the usage of a great deal of code and API calls. I learnt that VB6 can do
anything but it can also be hard work to make it do so. I could not have made 
this utility without the help of code from the various projects I have listed 
below.  

I hope you enjoy the functionality this utility provides. If you think 
you can improve anything then please feel free to do so. If you dislike my 
programming style then do keep those thoughts to yourself. :) 

Built on a 2.5ghz core2duo Dell Latitude E5400 running Windows 7 Ultimate 64bit 
using VB6 SP6.

   Tested on :
Windows 7 Pro 32bit on Intel
Windows 7 Ultimate 64bit on Intel
Windows XP SP3 on Intel
Windows 10 Home 64bit on AMD and Intel

 Dependencies:

 Notes:
Integers are retained (rather than longs) as some of these are passed to
library API functions in code that is not my own so I am loathe to change.

 Licence:
Copyright © 2019 Dean Beedell

This program is free software; you can redistribute it and/or modify it under 
the terms of the GNU General Public Licence as published by the Free Software 
Foundation; either version 2 of the License, or (at your option) any later 
version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY 
WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A 
PARTICULAR PURPOSE. See the GNU General Public Licence for more details.

You should have received a copy of the GNU General Public Licence along with 
this program; if not, write to the Free Software Foundation, Inc., 51 Franklin 
St, Fifth Floor, Boston, MA 02110-1301 USA

If you use this software in any way whatsoever then that implies acceptance of 
the licence. If you do not wish to comply with the licence terms then please 
remove the download, binary and source code from your systems immediately.

Credits:

I have really tried to maintain the credits as the project has progressed. If 
I have made a mistake and left someone out then do forgive me. I will make 
amends if anyone points out my mistake in leaving someone out.

Peacemaker2000    Original idea for a GDI+ dock came from here:
http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=55352&lngWId=1&fbclid=IwAR2FeR12CdaxyOoY-muw-b6_oDW-_19oLrt8syEL6BQSX4PMEfHyWpfqpzM

Olaf Schmidt - used some of Olafs code as examples of how to implement the 
handling of images using GDI+ and specifically used two routines, 
CreateScaledImg & ReadBytesFromFile.

Also critically, the idea of using the scripting dictionary as a repository for 
a collection of image bitmaps.

In addition, the easeing functions to do the bounce animation, I initially used 
a .js implementation but Olafs was better.

Spider Harper     Is64bit() function.

Wayne Phillips Used a heavily modified version of his code to bring an external 
application window to the foreground https://www.everythingaccess.com/tutorials.
asp?ID=Bring-an-external-application-window-to-the-foreground

www.thescarms.com Provided the code to enumerate through windows using a 
callback routine

dee-u Candon City, Ilocos Used a modified version of his code to obtain a window 
handle from a PID. https://www.vbforums.com/showthread.php?561413-getting-hwnd-
from-process

Shuja Ali @ codeguru for his settings.ini code.

An unknown, untraceable source, possibly on MSN - for the KillApp code

ALLAPI.COM        For the registry reading code.

Elroy on VB forums for his Persistent debug window
http://www.vbforums.com/member.php?234143-Elroy

Rxbagain on codeguru for his Open File common dialog code without a dependent OCX
http://forums.codeguru.com/member.php?92278-rxbagain

si_the_geek       for his special folder code

Aaron Young       for his code for registering a keypress system wide

                  Lots of GDI+ examples gleaned from here:
http://read.pudn.com/downloads29/sourcecode/windows/control/93919/Use_GDI+_(1627
568102003/frmMain.frm__.htm

La Volpe          Routine to check return value from any GDI++ function

Jacques Lebrun    Function to Provide resolution of shortcuts
https://www.vbforums.com/showthread.php?445574-Reading-shortcut-information

Dragokas systray code

Built using: VB6, MZ-TOOLS 3.0, CodeHelp Core IDE Extender Framework 2.2 & 
Rubberduck 2.4.1

*MZ-TOOLS https://www.mztools.com/
*CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=6246
8&lngWId=1
*Rubberduck http://rubberduckvba.com/
*Rocketdock https://punklabs.com/
*Registry code ALLAPI.COM
*http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWI
d=1
*PrivateExtractIcons code http://www.activevb.de/rubriken/
*Persistent debug code http://www.vbforums.com/member.php?234143-Elroy
*Open File common dialog code without dependent OCX-    
                http://forums.codeguru.com/member.php?92278-rxbagain
*Open font dialog code without dependent OCX

Background:

I always loved Rocketdock for its ease of use, the fact that it was free to 
download but more importantly because it allowed each of your PCs to display a 
desktop with a look and feel that was unique to you. At the time Objectdock and 
Rocketdock were developed. customisation was the name of the game, everyone was 
doing it and Rocketdock made it easy and painless to do so. Then came the change 
in the form of Windows 8 and 10... and a drastic rework to Microsoft's corporate 
policies meant that customisation was frowned upon, companies such as MS and 
Apple now wanted all your systems to look and operate just like everyone else's, 
corresponding to a corporate style. Each recent change to the Windows operating 
system has made it more closed and more difficult to customise. Slowly but 
surely, due to these changes being accepted by users, Rocketdock and other 
customisation tools fell out of favour and as a result Rocketdock's developers 
Punklabs moved onto better things and Rocketdock is no longer supported by them, 
meaning no more updates to fix bugs and no new versions with improved 
functionality. This is a problem for me as I want to use it on the latest 
versions of Windows.

Increasingly, Microsoft has changed Windows in unexpected and unpleasant ways 
introducing little problems for Rocketdock and its remaining users. Despite this, 
perhaps surprisingly, Rocketdock still works today even on Windows 10, however, 
it is becoming increasingly difficult to configure and operate with reliability. 
Without support and in the absence of new versions, users can struggle to make 
the old program run in the easy manner that they were accustomed to under 
Windows XP. With that in mind I stepped into the breach (as Rocketdock's self-
styled saviour) and I have created an open source version of something akin to 
RocketDock in both spirit and in design, this time named SteamyDock. Please note 
- SteamyDock hasn't used any of Rocketdock's code nor any of its resources, it 
has all been built from scratch.

I have been communicating with Skunkie from Punklabs and she has given me 
approval and encouragement to recreate Rocketdock in functionality and form, so 
that is what I have been doing. Instead of recreating Rocketdock as one 
monolithic tool I have instead decided to recreate it in three separate 
components, biting off what I can chew, as it were. First of all I have created 
the Icon Settings Tool, adding a lot more functionality than Rocketdock's 
original icon settings configuration tool. Secondly, I created the utility that 
replicates Rocketdock's Dock Settings configuration tool. Note that no code, nor 
resource has been taken from Punklabs nor Rocketdock, all the code, resources 
and icons shown are my own creation. All I have replicated is the functionality 
of Rocketdock and even then I have hopefully improved upon it and also upon the 
visual form.

SteamyDock is compatible with Rocketdock in many ways. It can use the same icons, 
the dock and icon configuration screens are very similar in operation so will be 
quite familiar to Rocketdock users. The main advantage of Steamydock over 
Rocketdock is that this new version is supported. In addition, it also has some 
new functionality and improvements. Fundamentally though, the design is limited 
to providing or enhancing what Rocketdock already provides. This will make the 
dock and its supporting utilities quite familiar to Rocketdock users.

This program will be available in two flavours. The first is a VB6 version, this 
is the original. The second will be a VB.NET version, not yet available. The two 
will be functionally the same, in almost all respects. The choice of which 
version to use will be entirely up to you. The VB.NET version has yet to be 
completed but when done will future-proof this utility. Regardless, the VB6 
version should work on Windows 10 for the 'foreseeable future' which means years 
and years yet to come as of 2021. Note that the VB6 version will also operate on 
ReactOS, a 32bit-only Windows clone when its WINE-inspired GDI+ layer is 
implemented

32bitness, I hear you say? - The VB6 version is of course 32 bit by default as 
that is all a VB6 application can ever be. VB6 is a 32bit language. Some see 
running 32bits as a limitation. It is not really, because of course, 32-bit 
applications run just fine on all versions of 64-bit Windows and this dock does 
not need 64bits to operate. This program has no need to use the main advantage 
of 64bit functionality - that being the ability to access more than 4gb of RAM. 
This utility does not require anywhere near the 4gigabytes maximum of RAM that 
32bit applications can address. In fact it averages just 45mb of usage even with 
seventy-two 128x128 bit icons displayed.

If you really do care about the 64bit thing and won't run 32bit programs on a 
64bit system for some personal reason, then run the .NET version when it comes 
out. Personally I prefer the 32bit VB6 version, I know them both inside out and 
VB6 is just 'better', easier to code and certainly much more fun to create.