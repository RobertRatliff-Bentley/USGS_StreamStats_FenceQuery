USGS StreamStats & VDOT Hydrology - Fence Query
Created by Robert Ratliff
Python v2.2.0 / VBA v1.0.1

INSTALL
  1. Close MicroStation.
  2. Double-click Install.bat.
  3. If asked, pick the MicroStation Macros folder to install into.

  The installer:
  - installs the Python libraries pyproj, shapely, openpyxl and numpy into MicroStation's Python.
    The bundled wheels are used, so no internet is needed. If they don't fit, it falls back to pypi.org.
  - copies USGS_StreamStats_FenceQuery.py and USGS_StreamStats_VDOT.mvba to the Macros folder.
    Any existing copies are backed up as *.bak-<date>.
  - checks that every library imports. The log is %TEMP%\USGS_StreamStats_Install.log.

RUN (MicroStation key-in)
  Python : python load "<Macros folder>\USGS_StreamStats_FenceQuery.py"
  VBA    : vba load USGS_StreamStats_VDOT
           vba run StreamStatsShow

OPTIONS
  Install.bat -MacrosDir "C:\path\Macros"   install into this folder without asking
  Install.bat -SkipLibraries                copy the tool files only
  Install.bat -PythonExe "C:\...\python.exe" use a different Python
  Install.bat -Quiet                        no prompts

UNINSTALL
  Double-click Uninstall.bat. This removes the tool files.
  The Python libraries are left in place because other tools may use them.

CONTENTS
  files\        the tool: Python script and VBA project
  wheels\       offline Python libraries (Windows x64, Python 3.12)
  vba_source\   exported VBA modules and forms, for reference
