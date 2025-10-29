This script downloads the latest weekly build of FreeCAD on Windows.\
\
1 - downloads the latest Windows installer to the Downloads directory\
  &nbsp;&nbsp;&nbsp;&nbsp;   If the latest is not yet online, it finds the latest.\
2 - extracts it to the specified portable software directory\
3 - makes a link to the bin file inside the portable software dirctory so you can use the shortcut in your taskbar.\
 &nbsp;&nbsp;&nbsp;&nbsp; the directory is opened in File Explorer so you can easily drag it to your taskbar.
\
Tested on Windows 10 Pro and Windows 11 Pro.

edit the portable software directory to suite your needs. I personally use a folder on my root directory named "Sortware-Portable"

~~~
# Variable for the portable directory (change this if needed)
$portableDir = "C:\Software-Portable"
~~~

You could run this script automatically every Wednesday and always be up to date on the bleeding edge of FreeCAD!
\
\
![Task Scheduler Image](https://i.imgur.com/q1CkHC8.png"TaskScheduler")
\
\
To all the FreeCAD developers and contributers, thank you for all your hard work. FreeCAD V1.1 is truly world class CAD software.
