import random

# This is a script which build a light and dark theme file that can
# be later used with AutoDark Mode to make custom day / night slideshows.

light_icon_data = """[Theme]
; Windows - IDS_THEME_DISPLAYNAME_LIGHT
DisplayName=@%SystemRoot%\System32\\themeui.dll,-2060
SetLogonBackground=0

; Computer - SHIDI_SERVER
[CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon]
DefaultValue=%SystemRoot%\System32\imageres.dll,-109

; UsersFiles - SHIDI_USERFILES
[CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon]
DefaultValue=%SystemRoot%\System32\imageres.dll,-123

; Network - SHIDI_MYNETWORK
[CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon]
DefaultValue=%SystemRoot%\System32\imageres.dll,-25

; Recycle Bin - SHIDI_RECYCLERFULL SHIDI_RECYCLER
[CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon]
Full=%SystemRoot%\System32\imageres.dll,-54
Empty=%SystemRoot%\System32\imageres.dll,-55

"""

dark_icon_data = """[Theme]
; Windows - IDS_THEME_DISPLAYNAME_DARK
DisplayName=@%SystemRoot%\System32\\themeui.dll,-2060
SetLogonBackground=0

; Computer - SHIDI_SERVER
[CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon]
DefaultValue=%SystemRoot%\System32\imageres.dll,-109

; UsersFiles - SHIDI_USERFILES
[CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon]
DefaultValue=%SystemRoot%\System32\imageres.dll,-123

; Network - SHIDI_MYNETWORK
[CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon]
DefaultValue=%SystemRoot%\System32\imageres.dll,-25

; Recycle Bin - SHIDI_RECYCLERFULL SHIDI_RECYCLER
[CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon]
Full=%SystemRoot%\System32\imageres.dll,-54
Empty=%SystemRoot%\System32\imageres.dll,-55

"""


general_cursor_data = """[Control Panel\Cursors]
AppStarting=%SystemRoot%\cursors\\aero_working.ani
Arrow=%SystemRoot%\cursors\\aero_arrow.cur
Crosshair=
Hand=%SystemRoot%\cursors\\aero_link.cur
Help=%SystemRoot%\cursors\\aero_helpsel.cur
IBeam=
No=%SystemRoot%\cursors\\aero_unavail.cur
NWPen=%SystemRoot%\cursors\\aero_pen.cur
SizeAll=%SystemRoot%\cursors\\aero_move.cur
SizeNESW=%SystemRoot%\cursors\\aero_nesw.cur
SizeNS=%SystemRoot%\cursors\\aero_ns.cur
SizeNWSE=%SystemRoot%\cursors\\aero_nwse.cur
SizeWE=%SystemRoot%\cursors\\aero_ew.cur
UpArrow=%SystemRoot%\cursors\\aero_up.cur
Wait=%SystemRoot%\cursors\\aero_busy.ani
DefaultValue=Windows Default
DefaultValue.MUI=@main.cpl,-1020

"""

general_desktop_data = """[Control Panel\Desktop]
Wallpaper=%SystemRoot%\web\wallpaper\Windows\img0.jpg
TileWallpaper=0
WallpaperStyle=10
Pattern=

"""

light_visual_styles_data = """[VisualStyles]
Path=%ResourceDir%\Themes\Aero\Aero.msstyles
ColorStyle=NormalColor
Size=NormalSize
AutoColorization=1
SystemMode=Light
AppMode=Light

"""

dark_visual_styles_data = """[VisualStyles]
Path=%ResourceDir%\Themes\Aero\Aero.msstyles
ColorStyle=NormalColor
Size=NormalSize
AutoColorization=1
SystemMode=Dark
AppMode=Dark

"""

general_screensaver_data = """[boot]
SCRNSAVE.EXE=

"""

general_master_theme_selector = """[MasterThemeSelector]
MTSM=RJSPBS

"""

general_sounds_data = """[Sounds]
; IDS_SCHEME_DEFAULT
SchemeName=@%SystemRoot%\System32\mmres.dll,-800

"""

# 1. Configuration

import os
import json

configuration_filename = 'configuration.json'
report_filename = 'report.json'
options = {}

if (os.path.exists(configuration_filename) and os.stat(configuration_filename).st_size != 0):
    # Load configuration file and previous settings
    print('Loading previous configuration.')
    with open('configuration.json') as configuration_file:
        options = json.load(configuration_file)
else:
    # First run, configure settings
    print('First run detected.')
    from tkinter.filedialog import askdirectory
    options['wallpaper directory'] = askdirectory(title = 'Select Wallpaper Folder')
    options['interval'] =  int(input("Select Interval between light wallpapers (in minutes): ")) * 60 * 60 * 1000
    if input("Include report of images analyzed? [Y/N]: ").upper == "Y":
        options['report'] = "True"
    else:
        options['report'] = "False"

    # Save settings to configuration file
    with open('configuration.json', 'w') as configuration_file:
        json.dump(options, configuration_file)


# 2. Analyze Wallpaper Brightness

import math
from PIL import Image, ImageStat

dark_both_threshold = 70  # 0 - 70 : dark, 70 - 80 : both
both_light_threshold = 80  # 70 - 80 : both, 80 - 180+ : light

wallpaper_info = {} # contains brightness of each wallpaper

for root, dirs, files in os.walk(options['wallpaper directory']):
    for name in files:
        if name.endswith('jpg'):
            image_path = os.path.join(options['wallpaper directory'], root, name)
            img = Image.open(image_path)
            try:
                stat = ImageStat.Stat(img)
                r, g, b = stat.mean
                brightness = round(math.sqrt(0.299*(r**2) + 0.587*(g**2) + 0.114*(b**2)), 2)
                print("%10s : %3.2f \t (%s)" % (name, brightness, ('dark' if brightness <= dark_both_threshold else ('both' if brightness <= both_light_threshold else 'light'))))
                wallpaper_info[image_path.replace('/', '\\')] = brightness
            except ValueError:
                print('Could not get wallpaper info of ' + name)
                pass


# 3. Categorize Wallpapers to Light, Dark

light = []
dark = []
for k, v in wallpaper_info.items():
    if v < dark_both_threshold:
        # wallpaper is used for dark theme only
        dark.append(k)
    elif v < both_light_threshold:
        # wallpaper is used both for light and dark theme
        light.append(k)
    else:
        # wallpaper is used for light theme only
        light.append(k)

random.shuffle(light)
random.shuffle(dark)

# 4. Wallpaper report (optional)

if bool(options['report']):
    with open(report_filename, 'w') as report_file:
        json.dump(wallpaper_info, report_file)

# 5. Build themes

# 5.1 Light Theme

f = open('Light.theme', 'w')
f.write(light_icon_data)
f.write(general_cursor_data)
f.write(general_desktop_data)
f.write(light_visual_styles_data)
f.write(general_screensaver_data)
f.write(general_master_theme_selector)
f.write(general_sounds_data)

general_slideshow_data = """[Slideshow]
Interval=""" + str(options['interval']) +  """
Shuffle=1
ImagesRootPath=""" + options['wallpaper directory'].replace('/', '\\') + "\n"

f.write(general_slideshow_data)

counter = 0
for i in light:
    f.write("Item" + str(counter) + "Path=" + i + '\n')
    counter += 1

f.close()

# 5.2 Dark Theme

f = open('Dark.theme', 'w')
f.write(dark_icon_data)
f.write(general_cursor_data)
f.write(general_desktop_data)
f.write(dark_visual_styles_data)
f.write(general_screensaver_data)
f.write(general_master_theme_selector)
f.write(general_sounds_data)
f.write(general_slideshow_data)

counter = 0
for i in dark:
    f.write("Item" + str(counter) + "Path=" + i + '\n')
    counter += 1

f.close()