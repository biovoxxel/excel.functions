from ij import IJ
from excel.functions.plugins import ExcelFunctions

imp = IJ.openImage("http://imagej.nih.gov/ij/images/blobs.gif");

imp.setRoi(36,30,63,54);
IJ.run(imp, "Measure", "");
imp.setRoi(116,32,63,54);
IJ.run(imp, "Measure", "");
imp.setRoi(62,99,63,54);
IJ.run(imp, "Measure", "");
imp.setRoi(164,113,63,54);
IJ.run(imp, "Measure", "");
imp.setRoi(20,175,63,54);
IJ.run(imp, "Measure", "");

ExcelFunctions.saveTableToWorkbookSheet("Results", "C:/Users/broch/Desktop/pythonTest.xlsx", "", 0, 0)
ExcelFunctions.addImageToWorkbookSheet(imp, "C:/Users/broch/Desktop/pythonTest.xlsx", 0, 5, 5)