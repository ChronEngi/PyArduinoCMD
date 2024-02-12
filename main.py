"""PyArduinoCMD. ChronEngi project on GitHub.
PyArduinoCMD is designed to be a small, portable program. It allows you to connect to an Arduino device of any kind and retrieve information from the computer.
For automation, you can create a .txt format script with a sequence of strings and pauses within it to execute.
You can also send strings individually or reset Arduino without interacting with the 'reset' button."""

import ctypes                   
import os                       
import serial.tools.list_ports
import serial
import sys
import time
from colorama import Fore

version = "1.0 Official 🪟"
commands_list = """(exit) Closes the program.

(list) Prints out the full commands list.

(clear) Cleans the output.

(reset) Restarts Arduino.

(boardinfo) Gets all the Arduino info.

(disconnect) Temporarily disconnects from the port.

(str.) Send a custom string to Arduino, example of use;
    > str.hello
    This will send the string "hello".

(strings) Used to execute a sequence of str. for greater automatization, example of use;
    > strings
    Lines number: 3
    1> str.hello!
    2> wait.4
    3> str.byee!
So what happens, in line 1> send "hello!", in line 2>, wait 4 seconds and in line 3> send "byee!".

(script.) It's essentially the exact same thing as strings.\nThe only difference is that instead of writing lines to the terminal\nyou prepare a file NAME.txt in the same folder as the program\nto execute a sequence of str. example of the use;
    > script.hello
    Reading hello.txt


(cmd.) Execute a Windows CMD command in the terminal, example of use;
    > cmd.color 4
    This will change terminal color to red.
"""

#   Get the avaitable ports
ports = serial.tools.list_ports.comports()
arduino =  serial.Serial()
portList = []
connected = False

# Send an error/info message
def msg(msg_title, msg_info, msg_type):

    # Generate the window
    ctypes.windll.user32.MessageBoxW(0, msg_info, msg_title, msg_type)
    
    # Set as first window
    ctypes.windll.user32.SetWindowPos(ctypes.windll.kernel32.GetConsoleWindow(), -1, 0, 0, 0, 0, 0x3 | 0x10)
    return

def connect(the_port):
    global arduino
    try:
        arduino = serial.Serial(port=the_port, baudrate=9600)
        print(f"PyArduinoCMD [{version}]\nBy ChronEngi\n" "Open Source: " + Fore.LIGHTBLUE_EX + "https://github.com/ChronEngi/PyArduinoCMD\n" + Fore.RESET)
        print("Type 'list' to get all commands.\n")
    except Exception as e:
        print(e)
        msg("Error!", f"Couldn't connect to the port: {the_port}", 0x2010)
        sys.exit()

# Terminal name
ctypes.windll.kernel32.SetConsoleTitleW('PyArduinoCMD')

# Clear terminal at start
os.system('cls')

while True:

    print(Fore.RESET + 'Scanning USB ports...\n')

    # Trying to connect to any USB
    while not connected:

        # Get the available ports
        ports = serial.tools.list_ports.comports()
        portList = []

        # Search available USB ports with Arduino
        for onePort in ports:
            portList.append(onePort.device)

        # Found a USB COM avaitable!
        if portList != []:
            msg("Found!", f"Found port: {portList[0]}\nPress OK to connect.", 0x2040)
            os.system("cls")
            connected = True
            connect(portList[0])

    while connected:
        user_input = input(Fore.LIGHTGREEN_EX + f"{portList[0]}> " + Fore.RESET)
    
        if user_input == "exit":
            sys.exit()

        if user_input == "list":
            print(commands_list)

        if user_input == "reset":
            arduino.setDTR(False)
            arduino.setDTR(True)
            print("Arduino has been restarted.\n")

        if user_input.startswith("str."):
            formatted_text = user_input[4:]
            arduino.write((formatted_text + r"\n").encode())
            print(f"Sent: {formatted_text}\n")
        
        if user_input.startswith("strings"):
            lines = int(input("Lines number: "))
            concatenated_lines = []

            print(Fore.LIGHTBLUE_EX + "--------------------" + Fore.LIGHTYELLOW_EX)

            for i in range(lines):
                line = input(Fore.LIGHTBLUE_EX + str(i + 1) + "> " + Fore.RESET)
                concatenated_lines.append(line)
            
            for i in range(lines):

                if concatenated_lines[i].startswith("wait."):
                    wait_time = float(concatenated_lines[i].split(".")[1])
                    # print(f"Waiting {wait_time} seconds in line {i + 1}")
                    time.sleep(wait_time)
                               
                else:
                    formatted_text = concatenated_lines[i][4:]
                    formatted_text += "\n"
                    arduino.write(formatted_text.encode())

            print(Fore.LIGHTBLUE_EX + "--------------------" + Fore.LIGHTYELLOW_EX)

            print("\n")

        if user_input.startswith("script."):
                file_name = user_input.split(".")[1] + ".txt"
                print(f"Searching for {file_name}")

                try:
                    with open(file_name, 'r') as file:
                        lines = file.readlines()
                    print(Fore.LIGHTBLUE_EX + "--------------------" + Fore.LIGHTYELLOW_EX)

                    for line in lines:
                        line = line.strip()

                        if line.startswith("wait."):
                            wait_time = float(line.split(".")[1])
                            time.sleep(wait_time)

                        else:
                            formatted_text = line[4:]
                            formatted_text += "\n"
                            arduino.write(formatted_text.encode())

                        print(f"{line}")
                    
                    print(Fore.LIGHTBLUE_EX + "--------------------" + Fore.RESET)

                except FileNotFoundError:
                    print(Fore.RED + "File not found.\nMake sure the name is correct and\nthat the file is in the same folder as the program." + Fore.RESET)

                print("\n")
        
        if user_input == ("clear"):
            os.system("cls")

        if user_input.startswith("cmd."):
            formatted_text = user_input[4:]
            print(f"Executing: {formatted_text}\n")
            os.system(formatted_text)

        if user_input.startswith("boardinfo"):
            print(f"Description: {onePort}\nSettings = {arduino.get_settings}\n")

        if user_input.startswith("disconnect"):
            arduino.close()
            os.system("cls")
            connected = False
            print("Disconnected.\n")

# Add a newline at the end of the file
