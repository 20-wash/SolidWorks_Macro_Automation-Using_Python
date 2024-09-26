import sys
import subprocess
import time
import os
import win32com.client  # For interacting with SolidWorks COM API
from write_macro_args import plate # Import the write_block_macro function
from writehole_macro_args import plate_w_hole  # Import the write_hole_macro function

def open_solidworks_and_run_macro(macro_file):
    """Opens SolidWorks and runs the specified macro."""
    # Assuming you have a command to open SolidWorks and run a macro AND Adjust this command based on your SolidWorks installation
    
    if len(sys.argv) != 6:
        print("Usage: python script.py <width1> <height1> <thickness> <width2> <height2>")
        sys.exit(1)

    solidworks_path = r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe"
    
# Start SolidWorks
    subprocess.Popen(solidworks_path)  # Open SolidWorks
    time.sleep(10)  # Wait for SolidWorks to fully load

    # Connect to SolidWorks COM API
    swApp = win32com.client.Dispatch("SldWorks.Application")
    swApp.Visible = True

    # Path to your macro
    macro_path = macro_file
    print(macro_path)

    # Extract the file name without extension
    macro_module = os.path.splitext(os.path.basename(macro_path))[0]+"1"

    print(macro_module)

    # Run the macro
    swApp.RunMacro(macro_path,macro_module, "main")


if __name__ == "__main__":
    # These can be modified based on user input or other sources
    width1 = float(sys.argv[1])   
    height1 = float(sys.argv[2])       
    thickness = float(sys.argv[3])     
    width2 = float(sys.argv[4])
    height2 = float(sys.argv[5])         

    # Step 1: Write the plate macro
    plate_macropath = plate(width1, height1, thickness)        # Make the name of the macro as it's objective

    # Step 2: Write the plate with holes macro
    plate_w_hole_macropath=plate_w_hole(width1, height1, thickness, width2, height2)

     # Step 3: Open SolidWorks and run the plate macro first, if needed
    open_solidworks_and_run_macro(plate_macropath)  # Run the plate macro

    # Step 4: Open SolidWorks and run the hole macro
    open_solidworks_and_run_macro(plate_w_hole_macropath)  # Run the hole macro
    

    