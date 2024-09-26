
def plate_w_hole(width1, height1, width2, height2, thickness):
    
    macro_content = f"""
 
    Dim swApp As Object

    Dim swModel As Object  ' Changed from ModelDoc2 to Object

    Dim swModelExt As Object  ' Changed from ModelDocExtension to Object

    Dim swPart As Object  ' Changed from SldWorks.PartDoc to Object

    Dim boolstatus As Boolean
    Dim longstatus As Long, longwarnings As Long


    Sub main()
        ' Initialize SolidWorks application
        Set swApp = Application.SldWorks

        ' Create a new part document
        Set swModel = swApp.NewDocument("C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2023\\templates\\Part.PRTDOT", 0, 0, 0)
        
        Set swModelExt = swModel.Extension
        
        
        ' Select the front plane for sketching
        swModelExt.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
        
        ' Start a new sketch on the front plane
        swModel.SketchManager.InsertSketch True
        
        ' Create a center rectangle for the block (width = 0.4m, height = 0.6m)
        Dim vSkLines As Variant
        vSkLines = swModel.SketchManager.CreateCenterRectangle(0, 0, 0, {width1/2}, {height1/2}, 0)
        
        ' Clear previous selections
        'swModel.ClearSelection2 True
        
        ' Create a smaller center rectangle for the hole (width = 0.05m, height = 0.1m)
        vSkLines = swModel.SketchManager.CreateCenterRectangle(0, 0, 0, {width2/2}, {height2/2}, 0)
        
        ' Set the named view to Trimetric and zoom to fit
        swModel.ShowNamedView2 "*Trimetric", 8
        swModel.ViewZoomtofit2
        
        ' Extrude the sketch to create a block with a hole (depth = 0.01m)
        Dim myFeature As Object
        Set myFeature = swModel.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, {thickness}, 0.01, False, False, False, False, 0,0, False, False, False, False, True, True, True, 0, 0, False)
        
        ' Disable contour selection
        swModel.SelectionManager.EnableContourSelection = False


        ' Save the part
        Dim savePath As String: savePath = "D:\\ANK\\My macros\\Blockwithhole\\Plate_w_hole.SLDPRT"
        swModel.SaveAs3 savePath, 0, 1 + 2
        
        ' Export the part
        Dim strnewpath As String
        strnewpath = Replace(savePath, ".SLDPRT", ".STL")
        swModel.SaveAs3 strnewpath, 0, 1 + 2
        
        ' Isometric view and zoom to fit    
        swModel.ShowNamedView2 "", 7
        swModel.ViewZoomtofit2
        
        ' Notify user
        MsgBox "Block created in sldprt and stl format and saved at " & savePath
    End Sub


    """

    macro_content = macro_content.format(width1=width1, height1=height1,width2=width2, height2=height2, thickness=thickness)

    # Step 2: Define the path where you want to save the macro file
    macro_path = r"D:\AnK\My macros\Blocks\plate_w_hole_Late.swb"
    
    with open(macro_path, 'w') as file:
        file.write(macro_content)
    
    print(f"Macro saved to {macro_path}")    

    return macro_path

# plate_w_hole(0.4,0.6,0.05,0.100,0.01)  # Need to do a function call to create a file;