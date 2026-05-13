run("Excel Macro Extensions");

dir = getDirectory("home") + "Desktop/Test.xlsx";

if (File.exists(dir)) {
	deleted = File.delete(dir);
}
run("Set Measurements...", "area mean standard modal min centroid center perimeter bounding fit shape feret's integrated median skewness kurtosis area_fraction display redirect=None decimal=3");
setOption("BlackBackground", true);
run("Blobs (25K)");
run("Make Binary");
blobImageID = getImageID();

run("Analyze Particles...", "display exclude clear summarize");
Plot.create("Plot_of_Results", "x", "Area");
Plot.add("Separated Bars", Table.getColumn("Area", "Results"));
Plot.setStyle(0, "blue,#a0a0ff,1.0,Separated Bars");
Plot.show();
plotImageID = getImageID();

Ext.xlsx_SaveAllTablesToWorkbook(dir, true);
Ext.xlsx_AppendTableAsRows("Results", dir, 0, false);
Ext.xlsx_SaveTableAsWorksheet("Results", dir, "WithSheetName", true);
Ext.xlsx_AddImageToWorksheet(blobImageID, dir, 0, 5, 5);
Ext.xlsx_AddExcelChartFromPlot(dir, 1, "AREA");
Ext.xlsx_AddImageToWorksheet(plotImageID, dir, 1, -1, -1);
Ext.xlsx_AppendArrayAsExcelColumn(newArray("Heading","b","c","d","e","f","g","h"), dir, 1, 0);
Ext.xlsx_AppendArrayAsExcelRow(newArray(0,1,2,3,4,5,6,7,8,9), dir, 1, 1);
close("*");
close("Results");
close("Summary");

workSheetNameString = Ext.xlsx_GetAllWorkbookSheetNames(dir);
workSheetNames = split(workSheetNameString, ",");
for (i = 0; i < workSheetNames.length; i++) {
	Ext.xlsx_ImportXlsxWorksheetAsTable(dir, workSheetNames[i], true);
}
Ext.xlsx_SetColumnDataFormat(dir, 2, 3, "@");
/*
run("Excel Macro Extensions");

dir = getDirectory("home") + "Desktop/Test.xlsx";
Ext.xlsx_ImportXlsxWorksheetAsTable(dir, "Results", true);
*/