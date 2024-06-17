# file2TimelineDaVinci
Script that makes each paragraph in a docx into Text+ clips with black background clips and adds them to timeline in DaVinci Resolve based on marker colors and position.

## Requirements:
- Python 3.12.3 (but other versions might work as I have tried it with 3.8 and 3.10)
- DaVinci Resolve 18.6 (for previous versions, you'll have to add the import for DaVinciResolve.py, check DaVinci Resolve API (https://resolvedevdoc.readthedocs.io/en/latest/readme_resolveapi.html))
- Import easygui (https://easygui.sourceforge.net/) and python-docx (https://python-docx.readthedocs.io/en/latest/)

*If you want to use other kind of file instead of docx, you will have to change the import, the reading file line and 
how the script manages the text inside it (the file.paragraphs and .text calls)*

## Install:
Simple, just make sure the Python that is being used by DaVinci has all requirements and drag "txt2video.py" to "C:\Users\Username\AppData\Roaming\Blackmagic Design\DaVinci Resolve\Support\Fusion\Scripts\Edit".
Then go into DaVinci, enter in Edit panel, go to Workspace in the upper bar, then go to Scripts and there you'll just have to click it to use it.

*Any problem with the script will be seeing in the Console, just follow the same path as before but click in Console instead of Scripts.*

## Use:
You will need to add marks to your timeline. Each mark signifies one paragraph. Blue ones are for lower screen text adn red ones for mid right text. Also for red marks, the script doesn't adjust everything,
you'll have to adjust manually the text and background.
![image](https://github.com/GabrielGutiP/file2TimelineDaVinci/assets/146023114/dea9a2ed-b61d-4e0a-9012-c6f5e4ab0148)

![image](https://github.com/GabrielGutiP/file2TimelineDaVinci/assets/146023114/f562874c-965b-455d-8f20-2a15cd6c8d27)
Example with Blue mark and Red mark with editing. Upper image is the Edit timeline and below the result.

Another thing to note is that text is parsed by an internal function I created to make paragraphs to fit my design. Feel free to change those too, as my code isn't exactly perfect for this task.

## Thanks to:
Snap Captions from Orson Lord, their amazing Lua code let me understand the DaVinci Resolve API.
