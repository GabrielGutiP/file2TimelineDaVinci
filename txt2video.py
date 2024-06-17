
import easygui
import docx

class currentDavinci(object):
    # Properties of all projects
    # resolve is defined in DaVinci by default
    projectManager = resolve.GetProjectManager()
    mediaStorage = resolve.GetMediaStorage()
    project = projectManager.GetCurrentProject()
    mediaPool = project.GetMediaPool()
    #timelineName = str(time()) # Requires: from time import time
    #os = os
    #sys = sys

# Get markers from DaVinci
project = currentDavinci.project
timeline = project.GetCurrentTimeline()
markers = timeline.GetMarkers()

# Select clips Text+
folder = None
try:
    # List of subfolder
    subfolder = currentDavinci.mediaPool.GetRootFolder().GetSubFolderList()
    for sub in subfolder:
        if sub.GetName() == "Bases":
            folder = sub
    
except:
    print("Folder not found, import bin with base clips")

clip_list = folder.GetClipList()
blue_clip = None
red_clip = None
green_clip = None
# List of markers
markers_frames = list(markers.keys())

# Clips for base text
for clip in clip_list:
    cn = clip.GetClipProperty("File Name")
    if cn == "txt-blue":
        blue_clip = clip
    elif cn == "txt-red":
        red_clip = clip
    elif cn == "txt-green":
        green_clip = clip
    elif cn == "under-box":
        under_clip = clip
    elif cn == "name-box":
        name_clip = clip
    elif cn == "option-box":
        option_clip = clip
                
# Generate paragraphs for DaVinci (text to format, color from the marker, multiplier indicates next line)
def paragraphGen(txt, color, multiplier = 1):
    # Max lenght for each type
    if color == "blue":
        l = 90
    elif color == "red":
        l = 44
    # Length of line * line number
    l = l*multiplier
    # To cut on max lenght
    cut = l
    # If len > to max character in screen, adds a \n in the closest space
    if len(txt) > l:
        accL = 0
        while txt[l-1] != " ":
            accL = accL + 1
            # Just check 20 characters back
            if accL == 20:
                txt = txt[:cut-1]+"\n"+txt[cut:]
                break
            l = l-1
        if accL != 20:
            # String are inmutable in Python. need to change to list then back to string
            aux = list(txt)
            aux[l-1] = "\n"
            txt = "".join(aux)
        # Recursive call to ensure all text is formated
        txt = paragraphGen(txt, color, multiplier+1)
    return txt

# Create Text+ in Track based of Marker color
def createText(txt, recordF, mf, track, nameT):
    # Ensure that start frame cannot be lower than 0
    startframe = 0
    endframe = mf - recordF
    if recordF == 0:
        recordFrame = 216000
    else:       
        recordFrame = 216000 + recordF

# mediaPoolItem: Base object that is copied to make the clip
# startFrame: 0 to initiate the clip from the beguinning
# endframe: Clip lenght - 1
# trackIndex: Track where the clip goes
# recordFrame: Position in timeline where the clip will start
# Timeline starts with an hour -> 01:00:00:00 // hh:mm:ss:ms
# The starting frame of the timeline is 1 hour in frames at 60 fps = 216000
         
    if markers[mf]["color"] == "Red":
        # Text clip
        subClip = {
                "mediaPoolItem": red_clip,
                "startFrame": startframe,
                "endFrame": endframe,
                "trackIndex": track,
                "recordFrame": recordFrame,
            }
        txt = paragraphGen(txt, "red")
        # Black background, track - 1 to be under the text clip
        blackClip = {
                "mediaPoolItem": option_clip,
                "startFrame": startframe,
                "endFrame": endframe,
                "trackIndex": track - 1,
                "recordFrame": recordFrame,
            }
        
    elif markers[mf]["color"] == "Blue" and nameT == False:
        subClip = {
                "mediaPoolItem": blue_clip,
                "startFrame": startframe,
                "endFrame": endframe,
                "trackIndex": track,
                "recordFrame": recordFrame,
            }
        txt = paragraphGen(txt, "blue")
        blackClip = {
                "mediaPoolItem": under_clip,
                "startFrame": startframe,
                "endFrame": endframe,
                "trackIndex": track - 1,
                "recordFrame": recordFrame,
            }
            
    # Names in text
    elif nameT == True:
        subClip = {
                "mediaPoolItem": green_clip,
                "startFrame": startframe,
                "endFrame": endframe,
                "trackIndex": track,
                "recordFrame": recordFrame,
            }
        blackClip = {
                "mediaPoolItem": name_clip,
                "startFrame": startframe,
                "endFrame": endframe,
                "trackIndex": track - 1,
                "recordFrame": recordFrame,
            }
        
    # Add clip to active timeline
    clipInTimeline = currentDavinci.mediaPool.AppendToTimeline([subClip])[0]
    # Steps to insert text in Text+
    comp = clipInTimeline.GetFusionCompByIndex(1)
    textplusTools = comp.GetToolList(False, "TextPlus")
    textplusTools[1].SetInput("StyledText", txt)
    # Add background clip to active timeline
    currentDavinci.mediaPool.AppendToTimeline([blackClip])[0]

# ------------------- Main Script -------------------

# Select path for .docx
path = easygui.fileopenbox()

# Extract paragraph and add it to porject in DaVinci
file = docx.Document(path)

acc = 0 # Accumulator for loop

# Track for names and background
timeline.AddTrack("video")
timeline.AddTrack("video")

# Track for body text and background
timeline.AddTrack("video")
timeline.AddTrack("video")

# Track for player option 1 and background
timeline.AddTrack("video")
timeline.AddTrack("video")

# List of each paragraph from the file
allText = file.paragraphs
# Accumulator to move in the list of paragraphs
acc = 0
recordF = 0
for mf in markers_frames:
    # Extract paragraph string
    txt = allText[acc].text
    # Get number of video tracks (Must be 7 at start)
    totalTracks = timeline.GetTrackCount("video")
    # RecordFrame accumulator to keep them starting where the last one ends
    if acc == 0:
        recordF = 0

    if ": " in txt:
        # Separate person talking from main text
        header, body = txt.split(": ")
        if ("Jugador" in header) and ("O" not in header):
            createText(body, recordF, mf, 7, False)
            acc = acc + 1
        elif "Jugador O1" in header:
            while "Jugador O" in header:
                totalTracks = timeline.GetTrackCount("video")

                if "Jugador O1" in header:
                    createText(body, recordF, mf, 7, False)
                elif "Jugador O2" in header:
                    # If there isn't track for this options, create it
                    if totalTracks < 9:
                        timeline.AddTrack("video")
                        timeline.AddTrack("video")
                    createText(body, recordF, mf, 9, False)
                elif "Jugador O3" in header:
                    # If there isn't track for this options, create it
                    if totalTracks < 11:
                        timeline.AddTrack("video")
                        timeline.AddTrack("video")
                    createText(body, recordF, mf, 11, False)
                elif "Jugador O4" in header:
                    # If there isn't track for this options, create it
                    if totalTracks < 13:
                        timeline.AddTrack("video")
                        timeline.AddTrack("video")
                    createText(body, recordF, mf, 13, False)
                # Prepare next cicle to check condition
                acc = acc + 1
                txt = allText[acc].text
                # Make sure text can be split, if not, break while
                if not ": " in txt:
                    break
                header, body = txt.split(": ")
        else:
            if "Ados" in header:
                header = "Jefe"
            createText(header, recordF, mf, 5, True)
            createText(body, recordF, mf, 3, False)
            acc = acc + 1
    else:
        createText(txt, recordF, mf, 3, False)
        acc = acc + 1
    
    # RecordFrame accumulator
    recordF = mf


# Test for console in DaVinci

# projectManager = resolve.GetProjectManager()
# mediaStorage = resolve.GetMediaStorage()
# project = projectManager.GetCurrentProject()
# mediaPool = project.GetMediaPool()
# root_folder = mediaPool.GetRootFolder()
# clip_list = root_folder.GetClipList()
# for clip in clip_list:
#     cn = clip.GetClipProperty("File Name")
#     if cn == "txt-blue":
#         blue_clip = clip
#     elif cn == "txt-red":
#         red_clip = clip
#     elif cn == "txt-green":
#         green_clip = clip
# # mediaPoolItem: Objeto base que se copia
# # startFrame: 0 para iniciar el clip desde el principio
# # endframe: Duración del clip - 1 por empezar en 0.
# # trackIndex: Pista en la que irá el clip
# # recordFrame: Posición en la timeline donde empieza el clip.
# # La timeline empieza con una hora -> 01:00:00:00 // hh:mm:ss:ms
# # Por lo que el inicio de la timeline es 1 hora en frames si vamos a 60 fps = 216000
# subClip = {
#         "mediaPoolItem": blue_clip,
#         "startFrame": 0,
#         "endFrame": 600,
#         "trackIndex": 2,
#         "recordFrame": 216000,
#     }
# mediaPool.AppendToTimeline([subClip])
    
