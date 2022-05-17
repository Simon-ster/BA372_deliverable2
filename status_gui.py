import tkinter as tk           #Importing the tkinter package as 'tk'
from tkinter import filedialog #Importing tkinter's filedialog class for picking files
import pyodbc

############################## event handlers #######################################

#Event handler for the 'Select text file' button click.
def get_text_file():  
  filename = filedialog.askopenfilename(title = "Select a text file") #Pop up the file browser;
                                                                      #Assign its result to 'filename'
  text_file_entry.delete(0, tk.END)   #Delete everything in the entry box; from position 0 to the end.
  text_file_entry.insert(0, filename) #Insert filename in the entry box at position 0.

#Event handler for the 'Close' button.
def close_it():
  exit(0) #exit with status 0 (all is OK).

#Event handler for the 'Display' button
def run_it():
  #Clear the display_textbox
  display_textbox.delete("1.0", tk.END)  #Delete everything in the textbox; from position 1.0 to the end.

  #Read text from the file listed in the entry box.
  my_file = text_file_entry.get()
  try:
    #fi = open(my_file, "r")
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=path where you stored the Access file\file name.accdb;')
    cursor = conn.cursor()
  except:
    display_textbox.insert("1.0", "Error Connecting to Database...")
    return()

  my_string = fi.read()
  display_textbox.insert("1.0", my_string) #Insert the read file contents into the display_textbox,
                                           #starting at position 1.0.

################################# main #########################################

#Create the root window.
window = tk.Tk()
window.title("Course Evaluation Checker")

#Three frames
text_file_frame = tk.Frame(master = window) #Frame for the file browser button and associated text entry box.
display_frame = tk.Frame(master = window)   #Frame for the large textbox in the center of the GUI.
run_close_frame = tk.Frame(master = window) #Frame for the 'Run' and 'Close' buttons.

#Pack the three frames from top to bottom, in order of packing.
text_file_frame.pack(side = tk.TOP, fill = tk.BOTH)
display_frame.pack(side = tk.TOP, fill = tk.BOTH)
run_close_frame.pack(side = tk.TOP)

#Button for selecting the text file: 
  #Assign to text_file_frame. 
  #When clicked, call get_text_file().
text_file_button = tk.Button(master = text_file_frame, text = "Enter Course_ID", command = get_text_file)

#Text entry for the text entry box:
  #Assign to text_file_frame. 
text_file_entry = tk.Entry(master = text_file_frame, width = 100)

#Pack both the button and the text entry from left to right.
text_file_button.pack(side = tk.LEFT)
text_file_entry.pack(side = tk.LEFT)

#Textbox for displaying things:
  #Assign to display_frame.
  #Pack to the left; make it fill the entire window.
display_textbox = tk.Text(master = display_frame, width = 113, height = 25)
display_textbox.pack(side = tk.LEFT, fill = tk.Y)

#Scrollbar for the display_textbox:
  #Assign to display_frame.
  #Pack to the left (but since packed after display_textbox, it is placed to the right of display_textbox.
scrollbar = tk.Scrollbar(master = display_frame)
scrollbar.pack(side = tk.LEFT, fill = tk.Y)

#Associate the scrollbar with the display_textbox: i.e., 
  #make the scrollbar scroll the textbox vertically.
display_textbox.config(yscrollcommand = scrollbar.set)

#Set the event handler -yview() called on display_textbox- for the scrollbar.
scrollbar.config(command = display_textbox.yview) 

#Two buttons; Display and Close:
  #Assign both to the run_close_frame.
  #When Display button is clicked, call run_it().
  #When Close button is clicked, call close_it().
run_button = tk.Button(text = "Display", master = run_close_frame, command = run_it)
close_button = tk.Button(text = "Close", master = run_close_frame, command = close_it)

#Pack both buttons, left to right.
run_button.pack(side = tk.LEFT)
close_button.pack(side = tk.LEFT)

#Start the main event manager.
window.mainloop()
