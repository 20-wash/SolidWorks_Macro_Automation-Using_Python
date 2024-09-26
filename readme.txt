					Running the Macros 

1.	Change the Path if necessary: 
solidworks_path = r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe"
	a.	final_automation_file.py

2.	In file: write_macro_args.py
	a.	Change the Version of Solidworks (2023/2022/2024). If necessary, change the path also.
Set swModel = swApp.NewDocument("C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2023\\templates\\Part.PRTDOT", 0, 0, 0)

	b.	You can simply create the folder in location. If you don’t have D drive, create folder anywhere and change the path 
Dim savePath As String: savePath = "D:\\ANK\\My macros\\Blocks.SLDPRT"
 
	c.	You can also change the path as above.
 macro_path = r"D:\AnK\My macros\Blocks\plate1.swb"


3.	In file: writehole_macro_args.py
	a.	Do the same changes as above
