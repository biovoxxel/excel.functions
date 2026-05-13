/**
 * 
 */
package excel.functions.plugins;

import java.io.File;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;

import excel.functions.utils.ExcelUtils;
import ij.ImagePlus;
import ij.gui.Plot;
import ij.measure.ResultsTable;

/**
 * @author BioVoxxel
 *
 */
public class ExcelFunctions {
	

	/**
	 * 
	 * @param ijResultsTable - instance of an ImageJ {@link ij.measure.ResultsTable}
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetName - desired name of the worksheet (if sheet name exists already, a new sheet will be created with appended numbering)
	 * @param includeColumnHeadings - set <i>true</i> to add table data headings or set <i>false</i> to skip headings when adding data to a sheet
	 */
	public static void saveTableToWorkbookSheet(ResultsTable ijResultsTable, File workbookFile, String sheetName, Boolean includeColumnHeadings) {			
		SaveTableAsWorksheet.saveTableAsWorkbookSheet(ijResultsTable, workbookFile, sheetName, includeColumnHeadings);
	}
	
	/**
	 * 
	 * @param ijResultsTable - instance of an ImageJ {@link ij.measure.ResultsTable}
	 * @param filePathToXlsxWorkbookFile - the String file path including the extension ".xlsx" to the Excel workbook file
	 * @param sheetName - desired name of the worksheet (if sheet name exists already, a new sheet will be created with appended numbering)
	 * @param includeColumnHeadings - set <i>true</i> to add table data headings or set <i>false</i> to skip headings when adding data to a sheet
	 */
	public static void saveTableToWorkbookSheet(ResultsTable ijResultsTable, String filePathToXlsxWorkbookFile, String sheetName, Boolean includeColumnHeadings) {
		File workbookFile = new File(filePathToXlsxWorkbookFile);
		SaveTableAsWorksheet.saveTableAsWorkbookSheet(ijResultsTable, workbookFile, sheetName, includeColumnHeadings);
	}
	
	
	/**
	 * Table column headings will be automatically included in the exported sheet and sheets will receive the title of the exported table as name.
	 * 
	 * @param ijResultsTable - instance of an ImageJ {@link ij.measure.ResultsTable}
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * 
	 */
	public static void saveTableToWorkbookSheet(ResultsTable ijResultsTable, File workbookFile) {
		
		String sheetName = ijResultsTable.getTitle();
		boolean includeColumnHeadings = true;
		
		SaveTableAsWorksheet.saveTableAsWorkbookSheet(ijResultsTable, workbookFile, sheetName, includeColumnHeadings);	
	}
	
	
	/**
	 * 
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param includeColumnHeadings - set <i>true</i> to add table data headings or set <i>false</i> to skip headings when adding data to a sheet
	 */
	public static void saveAllOpenTablesAsWorkbookSheets(File workbookFile, Boolean includeColumnHeadings) {
		SaveAllTablesToWorkbook.saveAllOpenTablesAsWorkbookSheets(workbookFile, includeColumnHeadings);
	}

	/**
	 * 
	 * @param filePathToXlsxWorkbookFile - the String file path including the extension ".xlsx" to the Excel workbook file
	 * @param includeColumnHeadings - set <i>true</i> to add table data headings or set <i>false</i> to skip headings when adding data to a sheet
	 */
	public static void saveAllOpenTablesAsWorkbookSheets(String filePathToXlsxWorkbookFile, Boolean includeColumnHeadings) {
		File workbookFile = new File(filePathToXlsxWorkbookFile);
		SaveAllTablesToWorkbook.saveAllOpenTablesAsWorkbookSheets(workbookFile, includeColumnHeadings);
	}
	
		
	/**
	 * This function adds an image open in ImageJ / Fiji as PNG to the defined worksheet in a defined Excel workbook
	 * <i>Note: the size of images added to Excel cannot be properly controlled via the Apache POI library. 
	 * Therefore, images normally appear distorted in Excel. Using the image properties in Excel and pressing "Reset" restores the correct aspect ratio.</i> 
	 * 
	 * @param imagePlus - any ImagePlus instance open in a Window in ImageJ
	 * @param workbookFile - The file instance which holds the Excel workbook
	 */
	public static void addImageToWorkbookSheet(ImagePlus imagePlus, File workbookFile) {
		
		Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		
		String imageName = imagePlus.getTitle();
		
		String sheetName = ExcelUtils.getUniqueSheetName(workbook, imageName);
		
		int column = 0;
		int row = 0;
		
		AddImageToWorksheet.addImageToWorksheet(imagePlus, workbookFile, sheetName, column, row);
	}
			
	
	/**
	 * This function adds an image open in ImageJ / Fiji as PNG to the defined worksheet in a defined Excel workbook
	 * <i>Note: the size of images added to Excel cannot be properly controlled via the Apache POI library. 
	 * Therefore, images normally appear distorted in Excel. Using the image properties in Excel and pressing "Reset" restores the correct aspect ratio.</i> 
	 * 
	 * @param imagePlus - any ImagePlus instance open in a Window in ImageJ
	 * @param workbookFilePath - file path to an Excel workbook file
	 * @param sheetName - name of the worksheet or the 0-based index of an existing sheet.
	 * @param column - 0-based column of the cell which should be the location of the upper left corner of the added image (negative values take the first empty column)
	 * @param row - 0-based row of the cell which should be the location of the upper left corner of the added image (negative values take the first empty row)
	 */
	public static void addImageToWorkbookSheet(ImagePlus imagePlus, String workbookFilePath, String sheetName, int column, int row) {
		
		File workbookFile = new File(workbookFilePath);
		
		AddImageToWorksheet.addImageToWorksheet(imagePlus, workbookFile, sheetName, column, row);
	}
	
	
	/**
	 * This function adds an image open in ImageJ / Fiji as PNG to the defined worksheet in a defined Excel workbook
	 * <i>Note: the size of images added to Excel cannot be properly controlled via the Apache POI library. 
	 * Therefore, images normally appear distorted in Excel. Using the image properties in Excel and pressing "Reset" restores the correct aspect ratio.</i> 
	 * 
	 * @param imagePlus - any ImagePlus instance open in a Window in ImageJ
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetName - name of the worksheet or the 0-based index of an existing sheet.
	 * @param column - 0-based column of the cell which should be the location of the upper left corner of the added image (negative values take the first empty column)
	 * @param row - 0-based row of the cell which should be the location of the upper left corner of the added image (negative values take the first empty row)
	 */
	public static void addImageToWorkbookSheet(ImagePlus imagePlus, File workbookFile, String sheetName, int column, int row) {
		AddImageToWorksheet.addImageToWorksheet(imagePlus, workbookFile, sheetName, column, row);
	}
	
	/**
	 * This function adds an image open in ImageJ / Fiji as PNG to the defined worksheet in a defined Excel workbook
	 * <i>Note: the size of images added to Excel cannot be properly controlled via the Apache POI library. 
	 * Therefore, images normally appear distorted in Excel. Using the image properties in Excel and pressing "Reset" restores the correct aspect ratio.</i> 
	 * 
	 * @param imagePlus - any ImagePlus instance open in a Window in ImageJ
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetIndex - 0-based index of an existing sheet.
	 * @param column - 0-based column of the cell which should be the location of the upper left corner of the added image (negative values take the first empty column)
	 * @param row - 0-based row of the cell which should be the location of the upper left corner of the added image (negative values take the first empty row)
	 */
	public static void addImageToWorkbookSheet(ImagePlus imagePlus, File workbookFile, int sheetIndex, int column, int row) {
		AddImageToWorksheet.addImageToWorksheet(imagePlus, workbookFile, ""+sheetIndex, column, row);
	}
	
	
	/**
	 * 
	 * @param array - Number[] array which should be added to the specified Row
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetName - name String of the worksheet (if not existing it will be created)
	 * @param startingColumn - 0-based starting column from which onwards the array data are added to the cells of the first empty row in the sheet
	 */
	public static void appendArrayAsExcelRow(Number[] array, File workbookFile, String sheetName, int startingColumn) {
				
			String[] stringArray = convertNumberArrayInStringArray(array);
			
			AppendArrayAsExcelRow.appendArrayAsTableRow(stringArray, workbookFile, sheetName, startingColumn);
	}

	/**
	 * 
	 * @param array - Number[] array which should be added to the specified Row
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetIndex - 0-based integer index of an existing sheet
	 * @param startingColumn - 0-based starting column from which onwards the array data are added to the cells of the first empty row in the sheet
	 */
	public static void appendArrayAsExcelRow(Number[] array, File workbookFile, int sheetIndex, int startingColumn) {
				
			String[] stringArray = convertNumberArrayInStringArray(array);
			
			AppendArrayAsExcelRow.appendArrayAsTableRow(stringArray, workbookFile, ""+sheetIndex, startingColumn);
	}

	
	/**
	 * 
	 * @param array - String[] array which should be added to the specified Row
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetName - name String of the worksheet (if not existing it will be created)
	 * @param startingColumn - 0-based starting column from which onwards the array data are added to the cells of the first empty row in the sheet
	 */
	public static void appendArrayAsExcelRow(String[] array, File workbookFile, String sheetName, int startingColumn) {
		AppendArrayAsExcelRow.appendArrayAsTableRow(array, workbookFile, sheetName, startingColumn);
	}
	
	/**
	 * 
	 * @param array - String[] array which should be added to the specified Row
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetIndex - 0-based integer index of an existing sheet
	 * @param startingColumn - 0-based starting column from which onwards the array data are added to the cells of the first empty row in the sheet
	 */
	public static void appendArrayAsExcelRow(String[] array, File workbookFile, int sheetIndex, int startingColumn) {
		AppendArrayAsExcelRow.appendArrayAsTableRow(array, workbookFile, ""+sheetIndex, startingColumn);
	}
	
	/**
	 * 
	 * @param array - Numbers[] array which should be added to the specified Row
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetName - name String of the worksheet (if not existing it will be created)
	 * @param startingRow - 0-based row at which the first entry of the provided array is added to the cells of the first empty column in the sheet
	 */
	public static void appendArrayAsExcelColumn(Number[] array, File workbookFile, String sheetName, int startingRow) {
		
		String[] stringArray = convertNumberArrayInStringArray(array);
		
		AppendArrayAsExcelColumn.appendArrayAsTableColumn(stringArray, workbookFile, sheetName, startingRow);
	}
	
	/**
	 * 
	 * @param array - Numbers[] array which should be added to the specified Row
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetIndex - 0-based integer index of an existing sheet
	 * @param startingRow - 0-based row at which the first entry of the provided array is added to the cells of the first empty column in the sheet
	 */
	public static void appendArrayAsExcelColumn(Number[] array, File workbookFile, int sheetIndex, int startingRow) {
		
		String[] stringArray = convertNumberArrayInStringArray(array);
		
		AppendArrayAsExcelColumn.appendArrayAsTableColumn(stringArray, workbookFile, ""+sheetIndex, startingRow);
	}
	
	/**
	 * 
	 * @param array - String[] array which should be added to the specified Row
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetName - name String of the worksheet (if not existing it will be created)
	 * @param startingRow - 0-based row at which the first entry of the provided array is added to the cells of the first empty column in the sheet
	 */
	public static void appendArrayAsExcelColumn(String[] array, File workbookFile, String sheetName, int startingRow) {
		AppendArrayAsExcelColumn.appendArrayAsTableColumn(array, workbookFile, sheetName, startingRow);
	}
	
	/**
	 * 
	 * @param array - String[] array which should be added to the specified Row
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetIndex - 0-based integer index of an existing sheet
	 * @param startingRow - 0-based row at which the first entry of the provided array is added to the cells of the first empty column in the sheet
	 */
	public static void appendArrayAsExcelColumn(String[] array, File workbookFile, int sheetIndex, int startingRow) {
		AppendArrayAsExcelColumn.appendArrayAsTableColumn(array, workbookFile, ""+sheetIndex, startingRow);
	}
	
	/**
	 * Allowed data formats can be found under:
	 * {@link org.apache.poi.ss.usermodel.BuiltinFormats}
	 * 
	 * 
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetName - name String of the worksheet
	 * @param columnNumber - the 0-based column number to which the format change should be applied
	 * @param dataFormatString - the actual format string as found in the BuiltinFormats
	 */
	public static void setColumnDataFormat(File workbookFile, String sheetName, int columnNumber, String dataFormatString) {
		SetColumnDataFormat.setColumnDataFormat(workbookFile, sheetName, columnNumber, dataFormatString);
	}
	
	/**
	 * 
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetIndex - 0-based index of an existing sheet
	 * @param columnNumber - the 0-based column number to which the format change should be applied
	 * @param dataFormatString - the actual format string as found in the BuiltinFormats
	 */
	public static void setColumnDataFormat(File workbookFile, int sheetIndex, int columnNumber, String dataFormatString) {
		SetColumnDataFormat.setColumnDataFormat(workbookFile, ""+sheetIndex, columnNumber, dataFormatString);
	}
	
	
	/**
	 * 
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetName - name String of the worksheet
	 * @param useFirstRowAsHeading - takes first sheet row as heading of the IJ table
	 */
	public static void importXlsxWorksheetAsTable(File workbookFile, String sheetName, boolean useFirstRowAsHeading) {
		ImportXlsxWorksheetAsTable.importXlsxWorksheetAsTable(workbookFile, sheetName, useFirstRowAsHeading);
	}
	
	/**
	 * 
	 * @param workbookFile - Instance of an Excel workbook file
	 * @param sheetIndex - 0-based index of an existing sheet
	 * @param useFirstRowAsHeading - takes first sheet row as heading of the IJ table
	 */
	public static void importXlsxWorksheetAsTable(File workbookFile, int sheetIndex, boolean useFirstRowAsHeading) {
		ImportXlsxWorksheetAsTable.importXlsxWorksheetAsTable(workbookFile, ""+sheetIndex, useFirstRowAsHeading);
	}
	
	
	/**
	 * 
	 * @param workbookFile - Instance of an Excel workbook file
	 * @return concatenated String (comma separated) of all existing worksheet names
	 */
	public static String getAllWorkbookSheetNames(File workbookFile) {
		return getAllWorkbookSheetNames(workbookFile);
	}
	
	
	/**
	 * Adds the specified ResultsTable as additional columns after the existing ones on the specified sheet
	 * 
	 * @param ijResultsTable - instance of an ImageJ {@link ij.measure.ResultsTable}
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetName - desired name of the worksheet (if sheet name exists already, a new sheet will be created with appended numbering)
	 * @param includeColumnHeadings - set <i>true</i> to add table data headings or set <i>false</i> to skip headings when adding data to a sheet
	 */
	public static void appendTableAsColumns(ResultsTable ijResultsTable, File workbookFile, String sheetName, Boolean includeColumnHeadings) {
		AppendTableAsColumns.appendTableAsColumns(ijResultsTable, workbookFile, sheetName, includeColumnHeadings);
	}
	
	
	/**
	 * Adds the specified ResultsTable as additional columns after the existing ones on the specified sheet
	 * 
	 * @param ijResultsTable - instance of an ImageJ {@link ij.measure.ResultsTable}
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetIndex - Integer index of desired worksheet (if name exists already, a new sheet will be created with appended numbering)
	 * @param includeColumnHeadings - set <i>true</i> to add table data headings or set <i>false</i> to skip headings when adding data to a sheet
	 */
	public static void appendTableAsColumns(ResultsTable ijResultsTable, File workbookFile, int sheetIndex, Boolean includeColumnHeadings) {
		AppendTableAsColumns.appendTableAsColumns(ijResultsTable, workbookFile, ""+sheetIndex, includeColumnHeadings);
	}
	
	
	/**
	 * Adds the specified ResultsTable as additional rows below the existing ones on the specified sheet
	 * 
	 * @param ijResultsTable - instance of an ImageJ {@link ij.measure.ResultsTable}
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetName - desired name of the worksheet (if sheet name exists already, a new sheet will be created with appended numbering)
	 * @param includeColumnHeadings - set <i>true</i> to add table data headings or set <i>false</i> to skip headings when adding data to a sheet
	 */
	public static void appendTableAsRows(ResultsTable ijResultsTable, File workbookFile, String sheetName, Boolean includeColumnHeadings) {
		AppendTableAsRows.appendTableAsRows(ijResultsTable, workbookFile, sheetName, includeColumnHeadings);
	}
	
	
	/**
	 * Adds the specified ResultsTable as additional rows below the existing ones on the specified sheet
	 * 
	 * @param ijResultsTable - instance of an ImageJ {@link ij.measure.ResultsTable}
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetIndex - Integer index of desired worksheet (if name exists already, a new sheet will be created with appended numbering)
	 * @param includeColumnHeadings - set <i>true</i> to add table data headings or set <i>false</i> to skip headings when adding data to a sheet
	 */
	public static void appendTableAsRows(ResultsTable ijResultsTable, File workbookFile, int sheetIndex, Boolean includeColumnHeadings) {
		AppendTableAsRows.appendTableAsRows(ijResultsTable, workbookFile, ""+sheetIndex, includeColumnHeadings);
	}
	
	
	/**
	 * 
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetName - desired name of an existing worksheet which should be deleted (does nothing if sheet is not existing)
	 */
	public static void deleteWorkbookSheet(File workbookFile, String sheetName) {
		DeleteWorkbookSheet.deleteWorkbookSheet(workbookFile, sheetName);
	}
	
	
	/**
	 * 
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetIndex - 0-based index of an existing sheet which should be deleted (does nothing if sheet is not existing)
	 */
	public static void deleteWorkbookSheet(File workbookFile, int sheetIndex) {
		DeleteWorkbookSheet.deleteWorkbookSheet(workbookFile, ""+sheetIndex);
	}
	
	
	/**
	 * 
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetName - desired name of an existing worksheet which should be deleted (does nothing if sheet is not existing)
	 * @param plot - Instance of an ImageJ Plot window
	 * @param chartType	- {@link org.apache.poi.xddf.usermodel.chart.ChartTypes}
	 */
	public static void addPlotToWorksheet(File workbookFile, String sheetName, Plot plot, ChartTypes chartType) {
		AddExcelChartFromPlot.addExcelChartFromPlot(workbookFile, sheetName, plot, chartType);
	}
	
	
	/**
	 * 
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetIndex - 0-based index of an existing sheet which should be deleted (does nothing if sheet is not existing)
	 * @param plot - Instance of an ImageJ Plot window
	 * @param chartType	- {@link org.apache.poi.xddf.usermodel.chart.ChartTypes}
	 */
	public static void addPlotToWorksheet(File workbookFile, int sheetIndex, Plot plot, ChartTypes chartType) {
		AddExcelChartFromPlot.addExcelChartFromPlot(workbookFile, ""+sheetIndex, plot, chartType);
	}
	
	
	/**
	 * 
	 * @param numberArray - a 2D array containing primitive numbers
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetName - desired name of an existing worksheet which should be deleted (does nothing if sheet is not existing)
	 */
	public static void add2DArrayToWorksheet(Number[][] numberArray, File workbookFile, String sheetName) {
		add2DArrayToWorksheet(numberArray, workbookFile, sheetName, 0, 0);
	}
	
	
	/**
	 * 
	 * @param numberArray - a 2D array containing primitive numbers
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetName - desired name of an existing worksheet which should be deleted (does nothing if sheet is not existing)
	 * @param startingRowIndex - the 0-based row index at which the data should be added
	 * @param startingColumnIndex - the 0-based column index at which the data should be added
	 */
	public static void add2DArrayToWorksheet(Number[][] numberArray, File workbookFile, String sheetName, int startingRowIndex, int startingColumnIndex) {
		
		String[][] stringArray = new String[numberArray.length][numberArray[0].length];
		for (int i = 0; i < numberArray.length; i++) {
			stringArray[i] = convertNumberArrayInStringArray(numberArray[i]);
		}
		
		add2DArrayToWorksheet(stringArray, workbookFile, sheetName, startingRowIndex, startingColumnIndex);
		
	}
	
	/**
	 * 
	 * @param stringArray - a 2D array containing Strings
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetName - desired name of an existing worksheet which should be deleted (does nothing if sheet is not existing)
	 */
	public static void add2DArrayToWorksheet(String[][] stringArray, File workbookFile, String sheetName) {
		add2DArrayToWorksheet(stringArray, workbookFile, sheetName, 0, 0);
	}
	
	
	/**
	 * 
	 * @param stringArray - a 2D array containing Strings
	 * @param workbookFile - The file instance which holds the Excel workbook
	 * @param sheetName - desired name of an existing worksheet which should be deleted (does nothing if sheet is not existing)
	 * @param startingRowIndex - the 0-based row index at which the data should be added
	 * @param startingColumnIndex - the 0-based column index at which the data should be added
	 */
	public static void add2DArrayToWorksheet(String[][] stringArray, File workbookFile, String sheetName, int startingRowIndex, int startingColumnIndex) {
		
		Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		
		Sheet sheet = ExcelUtils.getSheet(workbook, sheetName);
		
		ExcelUtils.write2DArrayToSheet(sheet, startingRowIndex, startingColumnIndex, stringArray);
		
		ExcelUtils.saveWorkbook(workbook, workbookFile);
		
		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	
	private static String[] convertNumberArrayInStringArray(Number[] array) {
		if (array != null) {
			String[] stringArray = new String[array.length];
			
			for (int index = 0; index < stringArray.length; index++) {
				stringArray[index] = array[index].toString();
			}
			return stringArray;			
		} else {
			return new String[0];
		}
	}
}
