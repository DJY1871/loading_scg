import subprocess
import time

command = [r"D:\test_auto_console\test.exe"]

textlist = []
proc = subprocess.Popen(command, stderr=subprocess.PIPE, stdout=subprocess.PIPE, shell=True, text=True)

while True:

    if proc.poll() is None:
        output, _ = proc.communicate()
        if output != '':
            textlist.append(output)
        continue

    if proc.poll() == 0 :
        print("Task completed")
        break
    elif proc.poll() < 0 :
        print("Error: {}".format(proc.poll()))
        break
    else:
        print("Task is running")
        


filename = "output.txt"

with open(filename, "w") as file:
    file.writelines(textlist)
