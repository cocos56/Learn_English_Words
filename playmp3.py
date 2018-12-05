import os
import subprocess

cwd = os.getcwd() + r'\Datas\Mp3'
print(cwd)
os.chdir(cwd)

cmd = 'ffplay -i absurd.mp3 -c:a 3 -af volume'
cmd = 'ffprobe.exe -i absurd.mp3'
#cmd = 'ffplay -i absurd.mp3 -ss 00:00:00 -t 00:02:05'

filename = 'absurd.mp3'

