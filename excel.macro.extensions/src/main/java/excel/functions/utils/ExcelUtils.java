/**
 * 
 */
package excel.functions.utils;

import java.awt.Dimension;
import java.awt.Window;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Arrays;
import java.util.List;
import java.util.logging.Logger;

import javax.imageio.ImageIO;
import javax.swing.JOptionPane;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.EmptyFileException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.DisplayBlanks;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import ij.ImagePlus;
import ij.WindowManager;
import ij.gui.Plot;
import ij.measure.ResultsTable;
import ij.text.TextPanel;
import ij.text.TextWindow;

/**
 * @author BioVoxxel
 *
 */
public class ExcelUtils {

	protected static final Logger logger = Logger.getLogger(ExcelUtils.class.getName());
	protected static final int EXCEL_SHEET_NAME_LIMIT = 30;
	
	
	public static boolean CLOSE_WORKBOOK_AFTER_SAVING = true;
//	private static final String FIRST_SHEET_INDICATOR = "FIRST_SHEET";
//	private static final String LAST_SHEET_INDICATOR = "LAST_SHEET";
//	
	

	
	
	public static ResultsTable getIJResultsTable(String ijTableName) {
		
		logger.finest("Requested table = " + ijTableName);
		ResultsTable table = null;
		
		if (ijTableName.equals("")) {
			
			table = getActiveTable();
			logger.finest("Active table = " + table);
		} else {
			
			table = ResultsTable.getResultsTable(ijTableName);
			logger.finest("Specified table = " + table);
		}
		return table;
	}	
	
	
	
	public static ResultsTable getActiveTable() {
		ResultsTable rt = null;
		Window win = WindowManager.getActiveTable();
		if (win!=null && (win instanceof TextWindow)) {
			TextPanel tp = ((TextWindow)win).getTextPanel();
			rt = tp.getOrCreateResultsTable();
		}
		return rt;
	}
	
	
	public static String[][] getIJTableAsRowColumn2DArray(ResultsTable ijResultsTable, boolean includeColumnHeadings) throws NullPointerException {
		
		logger.finest("Processing ResultsTable = " + ijResultsTable);
		if (ijResultsTable == null) throw new NullPointerException("ResultsTable = " + ijResultsTable);
		
		int rowCount = ijResultsTable.getCounter();
		String columnHeadingsAsString = ijResultsTable.getColumnHeadings();
		String[] columnHeadingsArray = columnHeadingsAsString.split("\t");
		
		if (includeColumnHeadings) {
			rowCount += 1;
		}
		
		String[][] tableArray = new String[rowCount][columnHeadingsArray.length];
		
		int startingRow = 0;
		if (includeColumnHeadings) {
			logger.finest("Including column headings");
			tableArray[0] = columnHeadingsArray;
			startingRow = 1;
		}
		logger.finest("Starting row = " + startingRow);
		
		for (int row = startingRow; row < rowCount; row++) {
			tableArray[row] = ijResultsTable.getRowAsString(row - startingRow).split("\t");
		}
		return tableArray;
	}
	
	public static String[][] trim2DArray(String[][] array, int rowStart, int rowEnd, int columnStart, int columnEnd) throws IllegalArgumentException {
		
		if (rowEnd > rowStart) throw new IllegalArgumentException("rowEnd ("+rowEnd+") > rowStart ("+rowStart+")");
		
		if (columnEnd > columnStart) throw new IllegalArgumentException("columnEnd ("+columnEnd+") > columnStart ("+columnStart+")");
				
		String[][] rowLimitedArray = Arrays.copyOfRange(array, rowStart, rowEnd);
		String[][] trimmedArray = new String[rowEnd-rowStart][columnEnd-columnStart];
		for (int rowIndex = 0; rowIndex < rowLimitedArray.length; rowIndex++) {
			trimmedArray[rowIndex] = Arrays.copyOfRange(rowLimitedArray[rowIndex], columnStart, columnEnd);
		}
		return trimmedArray;
	}
	
	
	public static byte[] getImageAsByteArray(ImagePlus imagePlus) throws IOException {
		byte [] imageData = null;
		
		ByteArrayOutputStream bos = new ByteArrayOutputStream();

		ImageIO.write(imagePlus.getBufferedImage(), "png", bos);
		imageData = bos.toByteArray();
		bos.close();
		logger.finest("Image byte[] array of length = " + imageData.length);
		
		return imageData;
	}
	
	
	public static String fixFilePath(String filePath) {
		
		if (filePath == null || filePath.equals("")) {
			filePath = System.getProperty("user.home") + File.separator + "NoFileNameGivenForTable.xlsx";
		} else {		
			filePath = addMissingFileExtension(filePath);
		}
		
		return filePath;
	}
	
	
	protected static String addMissingFileExtension(String filePath) {
		if (!filePath.endsWith(".xlsx")) {
			filePath = filePath + ".xlsx";
		}
		return filePath;
	}
	
	
	public static Workbook getWorkbook(String filePath) {	
		File outputFile = new File(filePath);
		return getWorkbook(outputFile);
	}
	
	
	public static Workbook getWorkbook(File workbookFile) {
		
		logger.finest("workbookFile=" + workbookFile);

		Workbook workbook = null;
		
		if(ExcelUtils.writableFile(workbookFile)) {
			
			try {
				workbook = WorkbookFactory.create(workbookFile);
				logger.finest("Workbook " + workbook + " loaded");
				
			} catch (EncryptedDocumentException | IOException | EmptyFileException e) {
				try {
					logger.finest("Create an empty workbook");
					workbook = WorkbookFactory.create(true);
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		} else {
			JOptionPane.showMessageDialog(null, "The desired output file cannot be accessed.\n"
					+ "Make sure it is not open in any other application and try again.");
			workbook = null;
		}
		
	
		logger.finest("returned workbook = " + workbook);
		return workbook;
	}
	
		
	public static Sheet getSheet(Workbook workbook, String sheetName)  {
		Sheet sheet = null;
		
		if (isNumeric(sheetName)) {
			int sheetIndex = Integer.parseInt(sheetName);
			sheet = workbook.getSheetAt(sheetIndex);
		} else {
			sheet = workbook.getSheet(sheetName);			
		}
		if (sheet == null) {
			sheet = workbook.createSheet(sheetName);
		}
		return sheet;
	}
	
	
	public static Sheet getSheet(Workbook workbook, int sheetIndex)  {
		Sheet sheet = null;
		
		try {
			sheet = workbook.getSheetAt(sheetIndex);
		} catch (IllegalArgumentException e) {
			//e.printStackTrace();
			String uniqueSheetName = getUniqueSheetName(workbook, "Sheet");
			sheet = workbook.createSheet(uniqueSheetName);
		}
		return sheet;
	}
	
	public static Sheet getUniqueSheet(Workbook workbook, String sheetName)  {
		if (isNumeric(sheetName)) {
			sheetName = "Sheet";
		}
		
		String uniqueSheetName = getUniqueSheetName(workbook, sheetName);
		Sheet sheet = workbook.createSheet(uniqueSheetName);
		return sheet;
	}
	
	
	public static String getUniqueSheetName(Workbook workbook, String sheetName) {
		
		String uniqueSheetName = sheetName;
					
		List<String> allSheetNames = getAllSheetNames(workbook);
		
		int uniqueSheetIndex = 1;
		
		while (allSheetNames.contains(uniqueSheetName)) {
			uniqueSheetName = sheetName + "-" + uniqueSheetIndex;
			uniqueSheetIndex++;
		}
		return uniqueSheetName;
	}
	
	
	public static List<String> getAllSheetNames(Workbook workbook) {
		
		int numberOfSheets = workbook.getNumberOfSheets();
		String[] allSheetNames = new String[numberOfSheets];
		
		for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {
			allSheetNames[sheetIndex] = workbook.getSheetName(sheetIndex);
		}
		
		List<String> listOfAllSheetNames = Arrays.asList(allSheetNames);
				
		return listOfAllSheetNames;
		
	}
	
	
	public static Sheet[] getAllSheets(Workbook workbook) {
		
		int numberOfSheets = workbook.getNumberOfSheets();
		Sheet[] allSheets = new Sheet[numberOfSheets];
		
		for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {
			allSheets[sheetIndex] = workbook.getSheetAt(sheetIndex);
		}
		
		return allSheets;
	}
	
	
	public static String[][] getSheetAsRowColumn2DArray(Sheet sheet) {
		
		int rowCount = getRowCount(sheet);
		int columnCount = getColumnCount(sheet);
	
		String[][] table2DArray = new String[rowCount][columnCount];
		
		for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
			Row currentRow = sheet.getRow(rowIndex);
			
			if (currentRow != null) {
				int lastCellNumber = currentRow.getLastCellNum();
				
				for (int columnIndex = 0; columnIndex < lastCellNumber; columnIndex++) {
					Cell currentCell = currentRow.getCell(columnIndex);
					logger.fine("Current cell = " + currentCell);

					if (currentCell != null) {

						CellType currentCellType = currentCell.getCellType();
						logger.fine("Cell type = " + currentCellType);
						
						switch (currentCellType) {
						case BLANK:
							table2DArray[rowIndex][columnIndex] = "";
							break;
						case NUMERIC:
							table2DArray[rowIndex][columnIndex] = "" + currentCell.getNumericCellValue();
							break;
						case STRING:
							table2DArray[rowIndex][columnIndex] = currentCell.getStringCellValue();
							break;
						case BOOLEAN:
							table2DArray[rowIndex][columnIndex] = Boolean.toString(currentCell.getBooleanCellValue());
							break;
						case ERROR:
							table2DArray[rowIndex][columnIndex] = ""; //Byte.toString(currentCell.getErrorCellValue());
							break;
						case FORMULA:
							table2DArray[rowIndex][columnIndex] = currentCell.getCellFormula();
							break;
						case _NONE:
							table2DArray[rowIndex][columnIndex] = "";
							break;
						}
					} else {
						table2DArray[rowIndex][columnIndex] = "";
					}
				}
			} else {
				for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
					table2DArray[rowIndex][columnIndex] = "";
				}
			}
		}
		return table2DArray;
	}
	
	
	public static int getRowCount(Sheet sheet) {
		int rowCount = sheet.getLastRowNum() + 1;
		logger.finest("getRowCount = " + rowCount);
		return rowCount;
	}
	
	
	public static int getColumnCount(Sheet sheet) {
		int lastRowIndex = sheet.getLastRowNum();
		
		int lastColumnIndex = 0;
		for (int rowIndex = 0; rowIndex < lastRowIndex; rowIndex++) {
			lastColumnIndex = Math.max(lastColumnIndex, sheet.getRow(rowIndex).getLastCellNum());
		}
		
		logger.finest("getColumnCount = " + lastColumnIndex);
		return lastColumnIndex;
	}
	
	public static ResultsTable createResultsTableFrom2DArray(String[][] table2DArray) {
		return createResultsTableFrom2DArray(table2DArray, false);
	}
	
	
	public static ResultsTable createResultsTableFrom2DArray(String[][] table2DArray, boolean useFirstRowAsHeading) {
		
		int rowCount = getArrayRowCount(table2DArray);
		int columnCount = getArrayColumnCount(table2DArray);
		int rowStartIndex = 0;
		if (useFirstRowAsHeading) {
			rowStartIndex = 1;
		}

		ResultsTable resultsTable = new ResultsTable(rowCount-rowStartIndex);

		for (int rowIndex = rowStartIndex; rowIndex < rowCount; rowIndex++) {
			for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {

				if (isNumeric(table2DArray[rowIndex][columnIndex])) {
					
					double currentCellValueAsDouble = Double.NaN;
					try {
						currentCellValueAsDouble = Double.parseDouble(table2DArray[rowIndex][columnIndex]);
					} catch (NumberFormatException e) {
						logger.warning(e.getMessage());
					}
					
					if (useFirstRowAsHeading) {
						resultsTable.setValue(table2DArray[0][columnIndex], rowIndex-rowStartIndex, currentCellValueAsDouble);
					} else {
						resultsTable.setValue(columnIndex, rowIndex-rowStartIndex, currentCellValueAsDouble);
					}
				} else {
					
					if (useFirstRowAsHeading) {
						resultsTable.setValue(table2DArray[0][columnIndex], rowIndex-rowStartIndex, table2DArray[rowIndex][columnIndex]);
					} else {
						resultsTable.setValue(columnIndex, rowIndex-rowStartIndex, table2DArray[rowIndex][columnIndex]);
					}	
				}
			}
		}
		return resultsTable;
	}
	
	
	public static int getArrayRowCount(String[][] table2DArray) {
		int arrayRowCount = table2DArray.length;
		logger.finest("arrayRowCount = " + arrayRowCount);
		return arrayRowCount;
	}
	
	
	public static int getArrayColumnCount(String[][] table2DArray) {
		int arrayColumnCount = 0;
		if(table2DArray.length > 0) {
			arrayColumnCount = table2DArray[0].length;			
		}
		logger.finest("arrayColumnCount = " + arrayColumnCount);
		return arrayColumnCount;
	}
	
	
	public static void write2DArrayToSheet(Sheet sheet, int startingRowIndex, int startingColumnIndex, String[][] table2DArray) {
		
		int rowCount = getArrayRowCount(table2DArray);
		int columnCount = getArrayColumnCount(table2DArray);
		
		for (int rowIndex = startingRowIndex; rowIndex < rowCount + startingRowIndex; rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (row == null) {
				row = sheet.createRow(rowIndex);
			}
			for (int columnIndex = startingColumnIndex; columnIndex < columnCount + startingColumnIndex; columnIndex++) {
				Cell cell = row.getCell(columnIndex);
				if (cell == null) {
					cell = row.createCell(columnIndex);
				}
				String cellValue = table2DArray[rowIndex-startingRowIndex][columnIndex-startingColumnIndex];
				
				if (isNumeric(cellValue)) {
					cell.setCellValue(Double.parseDouble(cellValue));
				} else {
					cell.setCellValue(cellValue);
				}
			}
		}		
	}
	
	
	public static void append2DArrayToSheetColumns(String[][] table2DArray, Sheet sheet) {
		
		int sheetColumnCount = getColumnCount(sheet);
		write2DArrayToSheet(sheet, 0, sheetColumnCount, table2DArray);
	}
	
	
	public static void append2DArrayToSheetRows(String[][] table2DArray, Sheet sheet) {
		int sheetRowCount = getRowCount(sheet);
		write2DArrayToSheet(sheet, sheetRowCount, 0, table2DArray);
	}
	
	
	public static String[][] convertInColumnArray(String[] array) {
		String[][] columnArray = new String[array.length][1];
		for (int rowIndex = 0; rowIndex < array.length; rowIndex++) {
			columnArray[rowIndex][0] = array[rowIndex];
		}
		return columnArray;
	}
	
	
	public static String[][] convertInRowArray(String[] array) {
		String[][] rowArray = new String[1][array.length];
		rowArray[0] = array;
		return rowArray;
	}
	
	
	public static void addExcelChartFromPlot(Sheet sheet, Plot plot, ChartTypes chartType) {
		
		float columnWidthInPixel = sheet.getColumnWidthInPixels(0);
		double rowHeightinPixel = (double) sheet.getDefaultRowHeightInPoints();
		Dimension plotDimension = plot.getSize();
		double plotWidth = plotDimension.getWidth();
		double plotHeight = plotDimension.getHeight();
		int chartWidthInColumns = (int) Math.round(plotWidth/columnWidthInPixel);
		int chartHeightInRows = (int) Math.round(plotHeight/rowHeightinPixel);
		
		String plotName = plot.getTitle();
		String xLabel = plot.getLabel("x".charAt(0));
		String yLabel = plot.getLabel("y".charAt(0));
		//String legendLabel = plot.getLabel("l".charAt(0));
		
		Float[] X_Values = ArrayUtils.toObject(plot.getXValues());
		Float[] Y_Values = ArrayUtils.toObject(plot.getYValues());
		
		XSSFSheet xssfSheet = (XSSFSheet)sheet;
		XSSFDrawing drawing = xssfSheet.createDrawingPatriarch();

		int startColumn = 0;
		int startRow = 0;
		XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, startColumn, startRow, startColumn + chartWidthInColumns, startRow + chartHeightInRows);
		XSSFChart chart = drawing.createChart(anchor);
		
		chart.setTitleText(plotName);
		chart.displayBlanksAs(DisplayBlanks.GAP);
		
		
		XDDFNumericalDataSource<Float> xValues = XDDFDataSourcesFactory.fromArray(X_Values);
		XDDFNumericalDataSource<Float> yValues = XDDFDataSourcesFactory.fromArray(Y_Values);
		
		XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
		categoryAxis.setTitle(xLabel);		
		
		XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
		valueAxis.setTitle(yLabel);
		
		XDDFChartData data = chart.createData(chartType, categoryAxis, valueAxis);
		data.addSeries(xValues, yValues);
		data.setVaryColors(false);
		
		chart.plot(data);
				
	}
	
	
	public static boolean isNumeric(Object object) {
		return isNumeric(object.toString());
	}

	
	@SuppressWarnings("unused")
	public static boolean isNumeric(String text) {
		if (text == null) {
			return false;
		}
		
		try {
		    double d = Double.parseDouble(text);
		} catch (NumberFormatException nfe) {
		    return false;
		}
		return true;
	}
	
	
	public static Boolean readBooleanMacroInput(Object parameter) {
		
		String param = parameter.toString();
		Boolean outputValue = false;
		
		if (param.equals("1") || param.equalsIgnoreCase("true")) {
			outputValue = true;
		} 
		
		return outputValue;
	}
	
	
	public static void setColumnDataFormat(Sheet sheet, Integer columnNumber, String dataFormatString) {
		
		if (dataFormatString.equals("")) {
			dataFormatString = "General";
		}
		
		Workbook workbook = sheet.getWorkbook();
		
		DataFormat dataFormat = workbook.createDataFormat();
    	
    	CellStyle cellStyle = workbook.createCellStyle();
    	
    	cellStyle.setDataFormat(dataFormat.getFormat(dataFormatString));
    	logger.finest("Used data format = " + dataFormat.getFormat(dataFormatString));
    	
    	int lastRowNumber = sheet.getLastRowNum();
    	logger.finest("Last row = " + lastRowNumber);
    	
    	for (int rowIndex = 0; rowIndex <= lastRowNumber; rowIndex++) {
			try {
				sheet.getRow(rowIndex).getCell(columnNumber).setCellStyle(cellStyle);
				logger.fine("Processing Row = " + rowIndex);				
			} catch (NullPointerException e) {
				logger.finest("NullPointerException at row = " + rowIndex);
			}
		}
	}
	
	
	public static void closeWorkbookAfterSaving(Boolean close) {
		CLOSE_WORKBOOK_AFTER_SAVING = close;
	}
	
	/**
	 * 
	 * Important!: test this before file is loaded or used by e.g. getWorkbook(workbookFile);
	 * As soon as the file is loaded into the system it might appear to not be writable anymore!
	 * 
	 * @param file
	 * @return true if file can be overwritten and is not blocked by another application
	 */
	public static boolean writableFile(File file) {
		//test is file is accessible or open in other application
		boolean writableFile = true;
		
		if (file.exists()) {
			
			try {
				
				File testFile = new File(file.getAbsolutePath());
				writableFile = file.renameTo(testFile);
				
			} catch (SecurityException e) {
				
				writableFile = false;
				e.printStackTrace();
				
			} catch (Exception e) {
				
				writableFile = true;
				e.printStackTrace();
				
			}
		}
		
		logger.finest("writableFile=" + writableFile);
		return writableFile;
	}
	
	
	public static boolean saveWorkbook(Workbook workbook, File outputFile) {
		
		boolean successfullySaved = false;
		
		FileOutputStream fos = null;
		try {
			
			Path temp = Files.createTempFile("workbook", ".xlsx");
						
			fos = new FileOutputStream(temp.toFile());
//			System.out.println(fos);
				
			workbook.write(fos);
			successfullySaved = true;
			
			logger.finest("Closing FileOutputStream " + fos);
			fos.close();
			
			logger.finest("Closing Workbook " + workbook);
			if (CLOSE_WORKBOOK_AFTER_SAVING) {
				workbook.close();
			}
			
			
			try {
				
				System.out.println("Atomic Move of temp file");
				Files.move(temp, outputFile.toPath(), StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
				
			} catch (Exception e) {
				
				System.out.println("Atomic Move not supported, running normal file replacement");
				
				Files.move(temp, outputFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
				
			}
				
		} catch (IOException e) {
			e.printStackTrace();
		} catch (EmptyFileException efe) {
			
			outputFile.delete();
			saveWorkbook(workbook, outputFile);
			//efe.printStackTrace();
		}
		
		return successfullySaved;
		
	}
	
	
	public static boolean saveWorkbook(Workbook workbook, String filePath) {
		try {
			File outputFile = new File(filePath);
			return saveWorkbook(workbook, outputFile);
			
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}
	
	
	public static void logMacroParameters(Object[] parameters) {
		for (int p = 0; p < parameters.length; p++) {
			logger.finest("Macro parameter " + p +" = " + parameters[p]);
		}
	}
	
	
	public static int getLastColumnNum(Sheet sheet) {
		int lastRowIndex = sheet.getLastRowNum();
		logger.finest("Last row = " + lastRowIndex);
		
		int lastColumnIndex = 0;
		for (int rowIndex = 0; rowIndex < lastRowIndex; rowIndex++) {
			lastColumnIndex = Math.max(lastColumnIndex, sheet.getRow(rowIndex).getLastCellNum());
		}
		
		logger.finest("Last column = " + lastColumnIndex);
		return lastColumnIndex;
	}
}
