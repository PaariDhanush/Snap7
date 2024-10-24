import snap7
from datetime import datetime
import tkinter as tk
import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import *
from PIL import ImageTk, Image
import threading #py -m pip install schedule
import time
import csv
import tkinter.font as font
import tkinter.messagebox as messagebox

# Define the PLC's IP address and rack/slot numbers
plc_ip = "192.168.0.2"  # Replace with your PLC's IP address
rack = 0
slot = 2  # Replace with your PLC's slot number

# Create a connection to the PLC
plc = snap7.client.Client()

# Create a list to store data with columns: S.NO, Index No, datetime
table_data = []

# Variable to track the current serial number, initialized to 0
serialNumber = 0


# Function to read and update data from the PLC
def read_plc_data():
    global serialNumber  # Use the global serialNumber variable
    while True:
        try:
            tryDateTimeNow = datetime.now()
            TryDateTimeNowTrim = tryDateTimeNow.strftime("%Y-%m-%d %H:%M:%S")
            tryTimeNow=TryDateTimeNowTrim[11:]
            #print(tryTimeNow)
            if tryTimeNow =="16:00:00":
                autoCSVexport()
                time.sleep(5) 
                clear_table()
                serialNumber = 0
            elif tryTimeNow =="00:20:00":
                autoCSVexport()
                time.sleep(5) 
                clear_table()
                serialNumber = 0    
                
            # Connect to the PLC
            plc.connect(plc_ip, rack, slot)

            
            data = plc.db_read(427, 900, 4)  # Adjust DB and offset as needed for your variable

            # Extract the bytes representing TIME_OF_DAY
            time_bytes = data[1::2]  # Select every other byte starting from the second byte

            # Convert the bytes to a human-readable string
            tod_value = ":".join(format(byte, '02x') for byte in time_bytes)

            # Extract the IndexNo (uncomment and adjust if needed)
            #for x in range(5):
            #    IndexNo=x
            IndexNo=tod_value[3:5]   
                # IndexNo=tod_value[3:5]

                    # Check if the IndexNo has incremented
            if len(table_data) == 0 or IndexNo != table_data[-1][1]:

                # Increment the serial number and append the data to the list with S.NO
                serialNumber += 1    
                if IndexNo == 0:
                    serialNumber = 0
                    # if index value is 0, then the serial no will start from 1,2,3....
                    serialNumber += 1
                    

                # Get the current date and time
                curdatetime = datetime.now()
                currentDateTime = curdatetime.strftime("%d-%m-%Y %H:%M:%S")
                delayTime=0
                if len(table_data) > 0:
                    last_row = table_data[-1]  # Access the last row
                    CloumnSNO, CloumnIndexNo, CloumnDatetime,CloumnDelayTime = last_row  # Assuming you have three columns
                    
                    #Time from of Table cloumn of Datetime for the last row
                    trimlastTime=CloumnDatetime
                    curLastTime=trimlastTime[11:]
                    #print(curLastTime)
                    
                    #Time from of curdatetime variable
                    trimNewTime=currentDateTime
                    curNewTime=trimNewTime[11:]
                    #print(curNewTime)
                    #----------------------------------
                    
                    time1 = datetime.strptime(curLastTime, "%H:%M:%S")
                    #print('Start time is :', time1.time())

                    time2 = datetime.strptime(curNewTime, "%H:%M:%S")
                    #print('End time is :', time2.time())

                    #difference time1 and 
                    TimeDiff = time2 - time1
                    #----------------------------------  
                    delayTime=TimeDiff.total_seconds() 
                    
                else:
                    delayTime=0

                table_data.append((serialNumber, IndexNo, currentDateTime,delayTime))

                # Update the tkinter table
                update_table()
                time.sleep(0) 
                
        except Exception as e:
            #print(f"Error communicating with PLC: {e}")
            ExcepDateTimeNow = datetime.now()
            ExcepDateTimeNowTrim = ExcepDateTimeNow.strftime("%Y-%m-%d %H:%M:%S")
            ExceptimeNow=ExcepDateTimeNowTrim[11:]
            #print(ExceptimeNow)
            if ExceptimeNow =="16:00:00":
                autoCSVexport()
                time.sleep(5) 
                clear_table()
                serialNumber = 0
            elif ExceptimeNow =="00:20:00":
                autoCSVexport()
                time.sleep(5) 
                clear_table()
                serialNumber = 0 
            

        finally:
            # Close the connection to the PLC
            plc.disconnect()

            # Wait for 60 seconds before reading again
            time.sleep(1)

# Function to update the tkinter table using Entry widgets
def update_table():
    # Clear the previous table
    for widget in table_frame.winfo_children():
        widget.destroy()
    
    # Create Entry widgets to represent the table
    for x in range(len(table_data) + 1):
        for y in range(len(columns)):
            if x == 0:
                label_text = columns[y]
            else:
                label_text = str(table_data[x - 1][y])
                
            entry = tk.Entry(
                master=table_frame,
                width=20,
                fg='black',
                font=('Times New Roman', 13, 'italic'),
            )
            entry.grid(row=x, column=y)
            entry.insert(tk.END, label_text)  # Use tk.END to insert at the end of the Entry

            # Configure the canvas scroll region
            canvas.update_idletasks()
            canvas.config(scrollregion=canvas.bbox("all"))
            
         

        
def autoCSVexport():            
    dateTimeNow = datetime.now()
    dateTimeNowTrim = dateTimeNow.strftime("%Y-%m-%d %H:%M:%S")
    trimTime=dateTimeNowTrim
    timeNow=trimTime[11:] 
    
    if "07:00:00" <= timeNow <= "16:10:00":
        shiftReport = datetime.now().strftime('DS_indexReport_%d-%m-%Y_%H-%M-%S.csv')
    else:
        shiftReport = datetime.now().strftime('NS_indexReport_%d-%m-%Y_%H-%M-%S.csv')
        
        
    # Specify the CSV file name and location
    #"C:/DailyIndexReport/"
    #filelocation="C:/DailyIndexReport/"
    filelocation="C:/DailyIndexReport/"
    
    filename = filelocation+shiftReport

    # Open the CSV file in write mode
    with open(filename, mode='w', newline='') as file:
        writer = csv.writer(file)
        
        # Write the table headers
        writer.writerow(columns)
        
        # Write the table data
        for row in table_data:
            writer.writerow(row)
    try:
        print(f"Data exported to {filename} successfully.")
        return True
    except Exception as e:
        print(f"Error exporting data: {e}")
        return False
        
            
# Define a function to clear all the items present in Treeview
def clear_table():
    # Clear the previous table data
    table_data.clear()
    
    # Clear the Entry widgets in the table frame
    for widget in table_frame.winfo_children():
        widget.destroy()

    # Configure the canvas scroll region
    canvas.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))    
    
def export_data_manually():
    time.sleep(0) 
    if autoCSVexport():
        
        show_success_popup()
    
def show_success_popup():
    messagebox.showinfo("Success", "Data exported successfully!")

# Create a tkinter window
root = tk.Tk()
root.config(background='white')
root.title('H Assembly Index Datalogger')
#root.geometry('600x400')
root.geometry('1900x1900')

image1 = Image.open("logo1.png")
test = ImageTk.PhotoImage(image1)

label1 = tkinter.Label(image=test)
label1.image = test

# Position image
label1.place(x='9', y='0.3')

#user_name = Label(root,text = "Assembly Indexing Report").place(x = 40,y = 60) 
var = StringVar()
myFont = font.Font(size=47)
label = Label( root, textvariable=var,bg='#00ffcc',fg='black', relief=RAISED )
label['font'] = myFont
var.set("Assembly Indexing Report")
label.pack()

# Create an export button
export_button = Button(root, text="Export Data", command=export_data_manually)
export_button.pack()

# Create a frame to hold the table and scrollbar
frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True)

# Create a canvas for the frame with the scrollbar
canvas = tk.Canvas(frame)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Create a scrollbar for the canvas
scrollbar = ttk.Scrollbar(frame, orient='vertical', command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
canvas.configure(yscrollcommand=scrollbar.set)

# Create a frame inside the canvas to hold the table
table_frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=table_frame, anchor=tk.NW)

# Define table columns
columns = ("S.NO", "Index No", "Date Time","Delay Time(sec)")


# Create a thread to read and update data from the PLC
thread = threading.Thread(target=read_plc_data)
thread.daemon = True
thread.start()

root.mainloop()  # Start the tkinter main loop
