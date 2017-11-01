import os
import re
import wave

myPath = "D:\\Audio"
extList = [".wav"]

libraryFile = open("soundlibrary.txt","w")

for root, dirs, files in os.walk(myPath):
    for filename in files:
        for extension in extList:
            if ( filename.lower().endswith(extension) and (not filename[0] == ".") ):
                fullpath = os.path.join(root, filename)
                duration = 0
                tags = ""
                print "Processing..." + fullpath
                try:
                	waveFile = wave.open(fullpath)
                	frames = waveFile.getnframes()
                	rate = waveFile.getframerate()
                	duration = frames / float(rate)
                	waveFile = open(fullpath)
                	for line in waveFile.readlines():
                		if ( "WAVE" in line ):
                			if ( "fmt" in line ):
                				end = line.index("fmt")
		                		cutTags = line[:end].replace("WAVE", "").replace("bext", "").replace("RIFF", "")
		                	elif ( "data" in line ):
		                		end = line.index("data")
		                		cutTags = line[:end].replace("WAVE", "").replace("bext", "").replace("RIFF", "")
		                	else:
		                		cutTags = line.replace("WAVE", "").replace("bext", "").replace("RIFF", "")
		                	tags = re.sub(r'\W+', ' ', cutTags)
		                	break
                	libraryFile.write(fullpath + "|" + str(duration) + "|" + extension + "|" + tags + "\n")
                except:
                	print "Error with file: " + fullpath