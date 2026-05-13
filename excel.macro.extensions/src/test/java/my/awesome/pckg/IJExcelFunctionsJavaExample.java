package my.awesome.pckg;
import java.io.File;

import org.apache.poi.xddf.usermodel.chart.ChartTypes;

import excel.functions.plugins.ExcelFunctions;
import ij.IJ;
import ij.ImagePlus;
import ij.Prefs;
import ij.WindowManager;
import ij.gui.Plot;
import ij.measure.Measurements;
import ij.measure.ResultsTable;
import ij.plugin.filter.ParticleAnalyzer;

public class IJExcelFunctionsJavaExample {
	
	private static ResultsTable RESULTS_TABLE = new ResultsTable();
	private static ResultsTable SUMMARY_TABLE = new ResultsTable();
    private static final boolean INCLUDE_HEADINGS = true;
    private static final int ROW_NUMBER = 5;
    private static final int COLUMN_NUMBER = 10;
    private static final int STARTING_COLUMN = 0;
    private static final int STARTING_ROW = 0;
    private static final String DATA_FORMAT_STRING = "@";  //The @ sign stands for "Text" formatting of cells
    private static File WORKBOOK_FILE = new File(IJ.getDirectory("home") + "Desktop/Test.xlsx");
    
    public static void main(String[] args) {
        
    	Prefs.blackBackground = true;
    	IJ.run("Blobs (25K)");
		IJ.run("Make Binary");
		ImagePlus imp = WindowManager.getCurrentImage();
		    	
    	ParticleAnalyzer.setSummaryTable(SUMMARY_TABLE);
		int CURRENT_PA_OPTIONS = ParticleAnalyzer.CLEAR_WORKSHEET|ParticleAnalyzer.RECORD_STARTS|ParticleAnalyzer.SHOW_MASKS|ParticleAnalyzer.SHOW_SUMMARY;
		int MEASUREMENT_PA_FLAGS = Measurements.AREA|Measurements.MEAN|Measurements.STD_DEV|Measurements.MODE|Measurements.MIN_MAX|Measurements.CENTROID|Measurements.CENTER_OF_MASS|Measurements.PERIMETER|Measurements.RECT|Measurements.ELLIPSE|Measurements.SHAPE_DESCRIPTORS|Measurements.FERET|Measurements.INTEGRATED_DENSITY|Measurements.MEDIAN|Measurements.SKEWNESS|Measurements.KURTOSIS|Measurements.AREA_FRACTION|Measurements.STACK_POSITION|Measurements.LIMIT|Measurements.LABELS;
		
		ParticleAnalyzer pa = new ParticleAnalyzer(CURRENT_PA_OPTIONS, MEASUREMENT_PA_FLAGS, RESULTS_TABLE, 0, Double.POSITIVE_INFINITY);
		pa.setHideOutputImage(true);
		pa.analyze(imp);
		
		RESULTS_TABLE.show("CompleteResults");
		SUMMARY_TABLE.show("Summary");
    	
        ExcelFunctions.saveAllOpenTablesAsWorkbookSheets(WORKBOOK_FILE, INCLUDE_HEADINGS);
        ExcelFunctions.addImageToWorkbookSheet(imp, WORKBOOK_FILE, 1, COLUMN_NUMBER, ROW_NUMBER);

        Plot plot = new Plot("TestPlot", "X", "Y");
        plot.add("circle", RESULTS_TABLE.getColumnAsDoubles(1), RESULTS_TABLE.getColumnAsDoubles(11));
        plot.show();
        plot.setStyle(0, "black,black,1,circle");	
        ExcelFunctions.addPlotToWorksheet(WORKBOOK_FILE, 1, plot, ChartTypes.AREA);
 
        //individual arrays can be added as columns or rows as follows
        String[] stringArray = new String[] {"a", "b", "c", "d", "e", "f", "g"};
		ExcelFunctions.appendArrayAsExcelColumn(stringArray, WORKBOOK_FILE, "NewSheet", STARTING_ROW);   
		
		Number[] numberArray = new Number[] {1.2f,2.565f,334.6d,3.0d,5,6,7,8,9};    //different primitive number formats will finally be converted to doubles in Excel
		ExcelFunctions.appendArrayAsExcelRow(numberArray, WORKBOOK_FILE, 2, STARTING_COLUMN);

        //specify the data format of specific columns
        ExcelFunctions.setColumnDataFormat(WORKBOOK_FILE, 0, 0, DATA_FORMAT_STRING);
        
    }
}

