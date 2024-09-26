# Function to write block macro dynamically
def plate(width, height, depth):
    macro_content = f"""
    Dim swApp As Object
    Dim swModel As ModelDoc2
    Dim swSketchMgr As SketchManager
    Dim swFeatureMgr As FeatureManager
    Dim swPart As PartDoc

    Sub main()
    ' Initialize SolidWorks application
    Set swApp = Application.SldWorks
    
    ' Create a new part document
    Set swModel = swApp.NewDocument("C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2023\\templates\\Part.PRTDOT", 0, 0, 0)
    Set swPart = swModel
    
    ' Select the front plane
    swModel.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    
    ' Start a new sketch
    Set swSketchMgr = swModel.SketchManager
    swSketchMgr.InsertSketch True
    
    ' Create a rectangle with specified dimensions
    Dim width As Double: width = {width} ' Width of the block in meters
    Dim height As Double: height = {height} ' Height of the block in meters
    
    swSketchMgr.CreateCenterRectangle 0, 0, 0, width / 2, height / 2, 0
    
    ' Exit the sketch
    swSketchMgr.InsertSketch True
    
    ' Extrude the sketch to create a block
    Set swFeatureMgr = swModel.FeatureManager
    Dim depth As Double: depth = {depth} ' Depth of the block in meters
    
    swFeatureMgr.FeatureExtrusion2 True, False, False, 0, 0, depth, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False
    
    ' Save the part
    Dim savePath As String: savePath = "D:\\ANK\\My macros\\Blocks.SLDPRT"
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


    macro_path = r"D:\AnK\My macros\Blocks\plate1.swb"
    
    with open(macro_path, 'w') as file:
        file.write(macro_content)

   
    return macro_path


# plate(0.200,0.080,0.020)