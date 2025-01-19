import tkinter as tk
from tkinter import messagebox
from functions import engine, constructPort
from config import *
# from oct2py import Oct2Py

import matlab.engine
from matlab_connection import run_macro_and_calculate_k_value, simulate_macro


def process_inputs(Start_Frequency, End_Frequency, Thickness, biggest_circle_radius, Number_of_layers, Scaling_Ratio, SubstrateLength, SubstrateWidth, groundThickness, SubstrateHeight, Dipole_Connection_Length, Dipole_Connection_Width, Feedline_length, Feedline_width, sToBeDefined, antennaMaterial, substrateMaterial, groundMaterial, Input_FileName_to_modify, Output_FileName_to_Save):
    # Replace this with your actual function logic
    engine(float(Start_Frequency), float(End_Frequency), float(Thickness), float(biggest_circle_radius), int(Number_of_layers), float(Scaling_Ratio), SubstrateLength, SubstrateWidth, groundThickness, SubstrateHeight, Dipole_Connection_Length, Dipole_Connection_Width, Feedline_length, Feedline_width, sToBeDefined, antennaMaterial, substrateMaterial, groundMaterial, Input_FileName_to_modify, Output_FileName_to_Save)
    # result = f"Processed inputs: {input1}, {input2}, {input3}, {input4}, {input5}"
    result = 'Design file for desired attributes has been saved!'
    return result


# Function triggered on button click
def on_submit():
    # Get values from text boxes
    Start_Frequency = entry1.get()
    End_Frequency = entry2.get()
    Thickness = entry3.get()
    biggest_circle_radius = entry4.get()
    Number_of_layers = entry5.get()
    Scaling_Ratio = entry6.get()
    SubstrateLength = entry7.get()
    SubstrateWidth = entry8.get()
    groundThickness = entry9.get()
    SubstrateHeight = entry10.get()
    Dipole_Connection_Length = entry11.get()
    Dipole_Connection_Width = entry12.get()
    Feedline_length = entry13.get()
    Feedline_width = entry14.get()
    sToBeDefined = entry15.get()
    antennaMaterial = entry16.get()
    substrateMaterial = entry17.get()
    groundMaterial = entry18.get()
    Input_FileName_to_modify = entry19.get()
    Output_FileName_to_Save = entry20.get()

    # Ensure no input is empty
    if not all([Start_Frequency, End_Frequency, Thickness, biggest_circle_radius, Number_of_layers, Scaling_Ratio, Dipole_Connection_Length, Dipole_Connection_Width, Feedline_length, Feedline_width, Input_FileName_to_modify, Output_FileName_to_Save]):
        messagebox.showerror("Input Error", "All fields must be filled!")
        return
    
    # Call the processing function and show the result
    result = process_inputs(Start_Frequency, End_Frequency, Thickness, biggest_circle_radius, Number_of_layers, Scaling_Ratio, float(SubstrateLength), float(SubstrateWidth), float(groundThickness), float(SubstrateHeight), float(Dipole_Connection_Length), float(Dipole_Connection_Width), float(Feedline_length), float(Feedline_width), sToBeDefined, antennaMaterial, substrateMaterial, groundMaterial, Input_FileName_to_modify, Output_FileName_to_Save)
    messagebox.showinfo("Result", result)




if GUI_MODE:
    # Create the main application window
    app = tk.Tk()
    app.title("FORA Antenna Design Tool")
    app.geometry("500x750")

    # Create input labels and entry boxes
    labels = ["Simulation start frequency in GHz:", "Simulation end frequency in GHz:", "Antenna surface thickness in mm:", "Biggest octagon circle radius in mm:", "Number of fractal layers:", "Fractal scaling ratio:", "Substrate length in mm:", "Substrate width in mm:", "Ground thickness in mm:", "Substrate thickness in mm:", "Dipole connection length in mm", "Dipole connection width", 'Feedline length:', 'Feedline Width:', "Frequency to check params: ", "Antenna material: ", "Substrate material: ", "Ground material: ", "Input macro fileName to modify:", "Output macro fileName to save:"]
    entries = []

    for i, label in enumerate(labels):
        tk.Label(app, text=label).grid(row=i, column=0, padx=10, pady=5, sticky=tk.W)
        entry = tk.Entry(app, width=30)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entries.append(entry)

    # Unpack entries into separate variables
    entry1, entry2, entry3, entry4, entry5, entry6, entry7, entry8, entry9, entry10, entry11, entry12, entry13, entry14, entry15, entry16, entry17, entry18, entry19, entry20 = entries

    # Add a Submit button
    submit_button = tk.Button(app, text="Submit", command = on_submit)
    submit_button.grid(row=len(labels), column=0, columnspan=2, pady=10)

    # Start the main event loop
    app.mainloop()

else:
    """
    in a loop, the parameters go into the function, the resulting output macro file goes under the macros/custom directory. a matlab file runs it for design and picking face. 
    then the k value for port extension coefficient is calculated and written in a file that is located at the working folder. this k value file read in the python file, the macro file is once again read and
    another macro code block added to the macro file, which is intended for the port design. the final macro is run with matlab once again and the results with corresponding parameters are written to the local
    directory.
    
    """
    #here add the for loop for the input parameters
    Output_FileName = 'output.mcs'
    Output_FileName_to_Save = 'macros/custom/' + Output_FileName
    
    
    
    Start_Frequency = 1
    End_Frequency = 5
    Thickness=0.035
    Scaling_Ratio = 0.91
    SubstrateLength = 60 #automatic
    SubstrateWidth = 60 #automatic
    groundThickness = 0.035  

    SubstrateHeight = 1.2
    Dipole_Connection_Length = 0.9 
    Dipole_Connection_Width = 0.5
    Feedline_length = 14 #automatic
    # Feedline_width = 0.3 
    sToBeDefined = 2.45
    antennaMaterial = 'Copper (annealed)'
    substrateMaterial = 'FR-4 (lossy)'
    groundMaterial = 'PEC'
    Feedline_width = Dipole_Connection_Length

    Input_FileName_to_modify = 'input' #template macro file in which the additional macro lines will be written.
    Output_FileName_to_Save = r'C:\Program Files (x86)\CST Studio Suite 2024\Library\Macros\custom\output' #This should point to 'CST Studio Suite 2024\Library\Macros\custom' with a macro file \output should be given to the name of the output macro file.
    Output_FileName = 'output'
    for biggest_circle_radius in [10, 12, 13, 15]: # circle radius in mm.
        for Number_of_layers in [3, 4, 5, 7]: ## multiple nested loops can be integrated in here to enable multiple parameter search.
            for Scaling_Ratio in [0.6, 0.7, 0.8, 0.91, 0.95]: ## multiple nested loops can be integrated in here to enable multiple parameter search.
                result = process_inputs(Start_Frequency , End_Frequency, Thickness, biggest_circle_radius, Number_of_layers, Scaling_Ratio, float(SubstrateLength), float(SubstrateWidth), float(groundThickness), float(SubstrateHeight), float(Dipole_Connection_Length), float(Dipole_Connection_Width), float(Feedline_length), float(Feedline_width), sToBeDefined, antennaMaterial, substrateMaterial, groundMaterial, Input_FileName_to_modify, Output_FileName_to_Save)
                #input_macro_path = should be under custom /without port
                if result:
                    isSuccess, k_value_file_path = run_macro_and_calculate_k_value(Output_FileName)
                    if isSuccess:
                        print('K value calculation is successful!')
                        k_coefficient = open(k_value_file_path, 'r').read()
                        print('k_coefficient: ', k_coefficient)
                        portShieldMaterial = 'PEC'
                        
                        isSuccess, new_filePath = constructPort(k_coefficient, portShieldMaterial, Output_FileName_to_Save)
                        if isSuccess:
                            print('port definition is performed in the macro file.')
                            newFileName = 'simulation'
                            isSuccess = simulate_macro(newFileName)
                        if isSuccess:
                            print('Simulation is successful!')
                        else:
                            print('Error in simulation!')


