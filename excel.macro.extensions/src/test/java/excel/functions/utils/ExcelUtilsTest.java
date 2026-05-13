package excel.functions.utils;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.fail;

import java.io.File;
import java.io.IOException;
import java.util.Vector;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import excel.functions.plugins.ExcelFunctions;
import ij.IJ;
import ij.ImagePlus;
import ij.Prefs;
import ij.WindowManager;
import ij.gui.Plot;
import ij.measure.Measurements;
import ij.measure.ResultsTable;
import ij.plugin.filter.ParticleAnalyzer;

class ExcelUtilsTest {
	
	
	private static ResultsTable RESULTS_TABLE = new ResultsTable();
	private static ResultsTable SUMMARY_TABLE = new ResultsTable();
	private static Plot MOCK_PLOT = null;
	private static ImagePlus BLOBS_IMAGE = null;
	private static File OUTPUT_FILE = null;
	private static Workbook MOCK_WORKBOOK;
	private static Sheet MOCK_SHEET;
	private static Row MOCK_ROW;
	private static Cell MOCK_CELL;
	private static final String MOCK_FILE_PATH = "C:/users/Admin/Desktop/text"; 

	@BeforeAll
	static void setUpBeforeClass() throws Exception {
		
		ExcelUtils.setupLogger();
		
		Prefs.blackBackground = true;
		IJ.run("Blobs (25K)");
		IJ.run("Make Binary");
		BLOBS_IMAGE = WindowManager.getCurrentImage();
		
		ParticleAnalyzer.setSummaryTable(SUMMARY_TABLE);
		//ParticleAnalyzer.setResultsTable(RESULTS_TABLE);
		int CURRENT_PA_OPTIONS = ParticleAnalyzer.CLEAR_WORKSHEET|ParticleAnalyzer.RECORD_STARTS|ParticleAnalyzer.SHOW_MASKS|ParticleAnalyzer.SHOW_SUMMARY;
		int MEASUREMENT_PA_FLAGS = Measurements.AREA|Measurements.MEAN|Measurements.STD_DEV|Measurements.MODE|Measurements.MIN_MAX|Measurements.CENTROID|Measurements.CENTER_OF_MASS|Measurements.PERIMETER|Measurements.RECT|Measurements.ELLIPSE|Measurements.SHAPE_DESCRIPTORS|Measurements.FERET|Measurements.INTEGRATED_DENSITY|Measurements.MEDIAN|Measurements.SKEWNESS|Measurements.KURTOSIS|Measurements.AREA_FRACTION|Measurements.STACK_POSITION|Measurements.LIMIT|Measurements.LABELS;
		
		ParticleAnalyzer pa = new ParticleAnalyzer(CURRENT_PA_OPTIONS, MEASUREMENT_PA_FLAGS, RESULTS_TABLE, 0, Double.POSITIVE_INFINITY);
		pa.setHideOutputImage(true);
		pa.analyze(BLOBS_IMAGE);
		
		RESULTS_TABLE.show("Complete Results");
		SUMMARY_TABLE.show("Summary");
		
	
		MOCK_PLOT = new Plot("TestPlot", "X", "Y");
		MOCK_PLOT.add("circle", RESULTS_TABLE.getColumnAsDoubles(1), RESULTS_TABLE.getColumnAsDoubles(11));
		MOCK_PLOT.show();
		MOCK_PLOT.setStyle(0, "black,black,1,circle");	
		
		OUTPUT_FILE = new File(System.getProperty("user.home") + "/Desktop/Test.xlsx");
		if (OUTPUT_FILE.exists()) {
			OUTPUT_FILE.delete();
		}
		
		MOCK_WORKBOOK = WorkbookFactory.create(true);
		MOCK_SHEET = MOCK_WORKBOOK.createSheet("One");
		MOCK_WORKBOOK.createSheet("Two");
		MOCK_ROW = MOCK_SHEET.createRow(0);
		MOCK_CELL = MOCK_ROW.createCell(0);
		MOCK_CELL.setCellValue(100);
		
	}

	@AfterAll
	static void tearDown() throws Exception {
		BLOBS_IMAGE.close();
		RESULTS_TABLE = null;
		SUMMARY_TABLE = null;
		MOCK_WORKBOOK.close();
		System.gc();
	}

	
	@Test
	void testGetLastColumnNum() {
		assertEquals(0, ExcelUtils.getLastColumnNum(MOCK_SHEET), "last column in mock table should be the first column with index 0");
	}

	@Test
	void testReadBooleanMacroInput() {
		assertTrue(ExcelUtils.readBooleanMacroInput(1));
		assertTrue(ExcelUtils.readBooleanMacroInput("1"));
		assertTrue(ExcelUtils.readBooleanMacroInput(true));
		assertTrue(ExcelUtils.readBooleanMacroInput("true"));
		
		assertFalse(ExcelUtils.readBooleanMacroInput("false"));
		assertFalse(ExcelUtils.readBooleanMacroInput(""));
		assertFalse(ExcelUtils.readBooleanMacroInput("0"));
		assertFalse(ExcelUtils.readBooleanMacroInput(0));
	}

	@Test
	void testSetColumnDataFormat() {

		ExcelUtils.setColumnDataFormat(MOCK_SHEET, 0, "0.00%");
		assertEquals(10, MOCK_CELL.getCellStyle().getDataFormat());
		ExcelUtils.setColumnDataFormat(MOCK_SHEET, 0, "@");
		assertEquals(CellType.NUMERIC, MOCK_CELL.getCellType());
		assertEquals(49, MOCK_CELL.getCellStyle().getDataFormat());
	}

	@Test
	void testAddMissingFileExtension() {
		String fixedFileExtension = ExcelUtils.addMissingFileExtension(MOCK_FILE_PATH);
		assertEquals(MOCK_FILE_PATH + ".xlsx", fixedFileExtension);
	}
	
	@Disabled
	void testCreateResultsTable() {
		fail("Not yet implemented");
	}

	@Test
	void testFixFilePath() {
		String fixedFilePath = ExcelUtils.fixFilePath("");
		assertEquals("C:\\Users\\Admin\\NoFileNameGivenForTable.xlsx", fixedFilePath);
		fixedFilePath = ExcelUtils.fixFilePath(MOCK_FILE_PATH);
		assertEquals(MOCK_FILE_PATH + ".xlsx", fixedFilePath);
	}

	@Test
	void testGetAllSheetNames() {
		Vector<String> expectedSheetNames = new Vector<String>();
		expectedSheetNames.add("One");
		expectedSheetNames.add("Two");
		assertEquals(expectedSheetNames, ExcelUtils.getAllSheetNames(MOCK_WORKBOOK));
	}

	@Test
	void testGetAllSheets() {
		assertEquals(MOCK_SHEET, ExcelUtils.getAllSheets(MOCK_WORKBOOK)[0]);
	}

	@Test
	void testGetIJResultsTable() {
		
		assertArrayEquals(RESULTS_TABLE.getColumn(1), ExcelUtils.getIJResultsTable("Complete Results").getColumn(1));
	}

	@Test
	void testGetImageAsByteArray() {
		
		int firstByteFromImageByteArray = 0;
		try {
			firstByteFromImageByteArray = ExcelUtils.getImageAsByteArray(BLOBS_IMAGE)[0];
		} catch (IOException e) {
			e.printStackTrace();
		}
		assertEquals(-119, firstByteFromImageByteArray);
	}

	

	@Test
	void testGetImageJTableRowsAsVector() {
		int row = 6;
		String[][] tableArray = ExcelUtils.getIJTableAsRowColumn2DArray(RESULTS_TABLE, false);
		String rowString = RESULTS_TABLE.getRowAsString(row);
		String[] rowArray = rowString.split("\t");
		
		assertArrayEquals(rowArray, tableArray[row]);
	}

	@Disabled
	void testGetSheet() {
		fail("Not yet implemented");
	}

		
	@Disabled
	void testGetSheetObjectAsString() {
		
	}

	@Test
	void testGetUniqueSheetName() {
		assertEquals("One-1", ExcelUtils.getUniqueSheetName(MOCK_WORKBOOK, "One"), "Is existing should add number 1");
		assertEquals("Three", ExcelUtils.getUniqueSheetName(MOCK_WORKBOOK, "Three"), "Is already unique");
	}

	@Disabled
	void testGetVectorContentSize() {
		
	}

	@Test
	void testGetWorkbookStringAlsoTestsFileAsInputParameter() {
	
		Workbook wb = ExcelUtils.getWorkbook("D:\\Programming\\Java\\git\\excel-table-macro-extensions\\excel.macro.extensions\\src\\test\\resources\\Test1.xlsx");
		assertEquals("Complete Results", wb.getSheetName(0));
		try {
			wb.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}

	
	@Test
	void testAddExcelChartFromPlot() {
		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.AREA);
		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.AREA3D);
		//ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.BAR);
//		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.BAR3D);
//		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.DOUGHNUT);
		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.LINE);
		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.LINE3D);
//		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.PIE);
//		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.PIE3D);
//		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.RADAR);
		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.SCATTER);
//		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.SURFACE);
//		ExcelUtils.addExcelChartFromPlot(MOCK_SHEET, MOCK_PLOT, ChartTypes.SURFACE3D);

	}
	
	
	@Test
	void testIsNumericObjectOrString() {
		assertTrue(ExcelUtils.isNumeric((Object)"100"), "Should be indicated as numeric (true)");
		assertFalse(ExcelUtils.isNumeric((Object)"Text"), "Should not be labelled numeric (false)");
	}

	
	
	@Test
	void testReadWorkbookStringAlsoTestsInputAsFile() {
		Workbook wb = ExcelUtils.getWorkbook("D:\\Programming\\Java\\git\\excel-table-macro-extensions\\excel.macro.extensions\\src\\test\\resources\\Test1.xlsx");
		Sheet sheet = wb.getSheet("Complete Results");
		
		assertEquals(RESULTS_TABLE.getCounter(), sheet.getLastRowNum());
		try {
			wb.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	@Disabled
	void testSaveWorkbookWorkbookFile() {
		fail("Not yet implemented");
	}

	@Test
	void testSaveWorkbookWorkbookString() {
		assertTrue(ExcelUtils.saveWorkbook(MOCK_WORKBOOK, System.getProperty("user.home") + File.separator + "SavingTest.xlsx"));
	}
	
	
	@Test
	void testWrite2DArrayToSheet() {
		String[][] stringArray = new String[][] {{"one","two","three"},{"four","five","six"},{"seven","eight","nine"}};
		//Number[][] numberArray = new Number[][] {{1,2,3},{4,5,6},{7,8,9}};
		
		Sheet sheet = MOCK_WORKBOOK.createSheet("2DArray");
			
		ExcelUtils.write2DArrayToSheet(sheet, 2, 3, stringArray);
		
		assertEquals("five", sheet.getRow(3).getCell(4).getStringCellValue());
	}
	
	
//	@Disabled("Method is private, convert first to public then run test")
//	void testConvertNumberArrayInStringArray() {
//		Number[] nullArray = null;
//		Number[] numberArray = new Number[]{10d, 3, 5, 4f};
//		String[] stringArrayOfNullInput = ExcelFunctions.convertNumberArrayInStringArray(nullArray);
//		String[] stringArrayOfNumberArrayInput = ExcelFunctions.convertNumberArrayInStringArray(numberArray);
//		
//		assertEquals(0, stringArrayOfNullInput.length);
//		assertEquals(4, stringArrayOfNumberArrayInput.length);
//	}

}
