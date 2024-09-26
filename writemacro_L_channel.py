def create_l_channel_macro(height, width, thickness,extrusion_length):
    # Create the macro content as a multi-line string
    macro_content = f""" 
    Dim swApp As Object
    Dim Part As Object
    Dim boolstatus As Boolean
    Dim longstatus As Long, longwarnings As Long

    Sub main()
        Set swApp = Application.SldWorks

        'New Document
        Set Part = swApp.NewDocument("C:\\ProgramData\\SolidWorks\\SOLIDWORKS 2023\\templates\\Part.prtdot", 0, 0, 0)
        Dim swPart As Object
        Set swPart = Part
        swApp.ActivateDoc2 "Part1", False, longstatus
        Set Part = swApp.ActiveDoc
        Part.ClearSelection2 True
        Part.SketchManager.InsertSketch True

        ' Creating the L-channel sketch
        ' Bottom line
        Set skSegment = Part.SketchManager.CreateLine(0, 0, 0, 0, {height}, 0)
        ' Right vertical line
        Set skSegment = Part.SketchManager.CreateLine(0, {height}, 0, {thickness}, {height}, 0)
        ' Top horizontal line
        Set skSegment = Part.SketchManager.CreateLine({thickness}, {height}, 0, {thickness}, {thickness}, 0)
        ' Left vertical line
        Set skSegment = Part.SketchManager.CreateLine({thickness}, {thickness}, 0, {width},  {thickness}, 0)
        ' Closing the L-channel by connecting back to the bottom line
        Set skSegment = Part.SketchManager.CreateLine({width},  {thickness}, 0, {width}, 0, 0)
        Set skSegment = Part.SketchManager.CreateLine({width}, 0, 0, 0, 0, 0)

        ' End Sketch
        Part.SketchManager.InsertSketch False

        ' Select the entire sketch for extrusion
        boolstatus = Part.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
    

        ' Extruding the L-channel
        Dim swFeature As Object
        Set swFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.5, 0, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
        Part.SelectionManager.EnableContourSelection = False
                
        'Clear selection
        Part.ClearSelection2 True
    End Sub
    """

    filename = r"D:\AnK\My macros\Blocks\L_channel_5.swb"

    # Write the macro content to a .swb file
    with open(filename, "w") as file:
        file.write(macro_content)

    print(f"Macro saved as {filename}")


# Example usage
# create_l_channel_macro(height=0.058889, width=0.05756, thickness=0.010331,extrusion_length=0.1)

create_l_channel_macro(0.08,0.08,0.02,0.5)