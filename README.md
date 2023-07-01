# SteamyDock

A VB6 GDI+ WoW64 dock for Reactos, XP, Win7, 8 and 10.

SteamyDock is a functional reproduction of the dock we all know and
love - Rocketdock for Windows from Punklabs - which in turn was a
clone of the Mac OS/X dock. Back in the Noughties, there were other
docks too such as ObjectDock, all of them good, some commercial, some
free. Our new dock, SteamyDock is also free, each allow you to create
your own dock using your own personal style and any icons you choose
to import from any location.

SteamyDock gets its name from the bundling of my own three dock
utilities with my own self-created Steampunk icon sets, ie. SteamyDock
is a dock and the icons are steamy... so there you have it.

BUILD: The program runs without any Microsoft plugins.

Built using: VB6, MZ-TOOLS 3.0, VBAdvance, CodeHelp Core IDE Extender
Framework 2.2 & Rubberduck 2.4.1

Links:

MZ-TOOLS https://www.mztools.com/  
CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
Rubberduck http://rubberduckvba.com/  
Rocketdock https://punklabs.com/  
Registry code ALLAPI.COM  
La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
PrivateExtractIcons code http://www.activevb.de/rubriken/  
Persistent debug code http://www.vbforums.com/member.php?234143-Elroy  
Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  
VBAdvance  
Fafalone for the enumerate Explorer windows code

LICENCE AGREEMENTS:

Copyright 2023 Dean Beedell

In addition to the GNU General Public Licence please be aware that you may use
any of my own imagery in your own creations but commercially only with my
permission. In all other non-commercial cases I require a credit to the
original artist using my name or one of my pseudonyms and a link to my site.
With regard to the commercial use of incorporated images, permission and a
licence would need to be obtained from the original owner and creator, ie. me.

Tested on :

ReactOS 0.4.14 32bit on virtualBox  
Windows 7 Professional 32bit on Intel  
Windows 7 Ultimate 64bit on Intel  
Windows 7 Professional 64bit on Intel  
Windows XP SP3 32bit on Intel  
Windows 10 Home 64bit on Intel  
Windows 10 Home 64bit on AMD  
Windows 11 64bit on Intel

Dependencies:

Requires a steamydock folder in C:\Users\<user>\AppData\Roaming\
eg: C:\Users\<user>\AppData\Roaming\steamydock
Requires a docksettings.ini file to exist in C:\Users\<user>\AppData\Roaming\PzEasteamydockrth
The above will be created automatically by the compiled program when run for the
first time.

GDI+
A windows-alike o/s such as Windows 7-11 or ReactOS.
OLEEXP.TLB placed in sysWoW64 - required to obtain the explorer paths only
during development. OLEEXP.TLB placed in sysWoW64 - required to obtain the
explorer paths.

oleexp.tlb should typically be located in SysWow64 (or System32 on a 32-bit
Windows install). You can register it manually using regtlib.exe on Win 7-10
systems or the newer utility on Win 11.

However, it should be sufficient to let VB6 register it for you. When you first
try to run or compile it will come up with the project references utility. Point
OLEEXP to the correct location (SysWoW64). You should only have one copy
installed. Only needed during development as the types are compiled in. Once
your project is compiled, the TLB is no longer used. It does not need to be
present on end user machines.

From the command line, copy the tlb to a central location (system32 or wow64
folder) and register it.

COPY OLEEXP.TLB %SystemRoot%\System32\
REGTLIB %SystemRoot%\System32\OLEEXP.TLB

In the VB6 IDE - project - references - browse - select the OLEEXP.tlb

Project References:
VisualBasic for Applications  
VisualBasic Runtime Objects and Procedures  
VisualBasic Objects and Procedures  
OLE Automation - drag and drop  
Microsoft Shell Controls and Automation  
Microsoft scripting runtime - for the scripting dictionary usage  
OLEEXP Modern Shell Interfaces for VB6, v5.1

Credits

I have really tried to maintain the credits as the project has progressed. If I
have made a mistake and left someone out then do forgive me. I will make amends
if anyone points out my mistake in leaving someone out.

MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and
early 2000s. Thanks for VB6.

Peacemaker2000 Original idea for a GDI+ dock came from here:
http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=55352&lngWId=1&fbclid=IwAR2FeR12CdaxyOoY-muw-b6_oDW-_19oLrt8syEL6BQSX4PMEfHyWpfqpzM

Olaf Schmidt - used some of Olafs code as examples of how to implement the
handling of images using GDI+ and specifically used two routines,
createScaledImg & ReadBytesFromFile.

Also critically, the idea of using the scripting dictionary as a repository
for a collection of image bitmaps.

In addition, the easeing functions to do the bounce animation, I initially
used a converted .js implementation but Olafs was better.

Spider Harper Is64bit() function.

Wayne Phillips Used a heavily modified version of his code to bring an external
application window to the foreground
https://www.everythingaccess.com/tutorials.asp?ID=Bring-an-external-application-window-to-the-foreground

www.thescarms.com Provided the code to enumerate through windows using a
callback routine

dee-u Candon City, Ilocos Used a modified version of his code to obtain a window
handle from a PID.
https://www.vbforums.com/showthread.php?561413-getting-hwnd-from-process

Shuja Ali @ codeguru for his settings.ini code.

An unknown, untraceable source, possibly on MSN - for the KillApp code

ALLAPI.COM For the registry reading code.

Elroy on VB forums for his Persistent debug window - no longer used but thanks
anyway!
http://www.vbforums.com/member.php?234143-Elroy

Rxbagain on codeguru for his Open File common dialog code without a dependent
OCX http://forums.codeguru.com/member.php?92278-rxbagain

si_the_geek for his special folder code

Aaron Young for his code for registering a keypress system wide

Lots of GDI+ examples gleaned from here:

http://read.pudn.com/downloads29/sourcecode/windows/control/93919/Use_GDI+_(1627568102003/frmMain.frm__.htm

La Volpe Routine to check return value from any GDI++ function

Jacques Lebrun Function to Provide resolution of shortcuts
https://www.vbforums.com/showthread.php?445574-Reading-shortcut-information

Fafalone for the enumerate Explorer windows code:
https://www.vbforums.com/showthread.php?818959-VB6-Get-extended-details-about-Explorer-windows-by-getting-their-IFolderView

Dragokas systray code
