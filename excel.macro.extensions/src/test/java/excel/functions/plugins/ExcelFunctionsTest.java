package excel.functions.plugins;

import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.fail;

import java.io.File;

import org.junit.jupiter.api.Disabled;

import ij.IJ;
import ij.ImagePlus;
import ij.Prefs;
import ij.WindowManager;
import ij.measure.Measurements;
import ij.measure.ResultsTable;
import ij.plugin.filter.ParticleAnalyzer;

class ExcelFunctionsTest {
	
	private static ResultsTable RESULTS_TABLE = new ResultsTable();
	private static ResultsTable SUMMARY_TABLE = new ResultsTable();
	private static ImagePlus BLOBS_IMAGE = null;
	private static File OUTPUT_FILE = null;
	private static final String COLUMN_HEADING = "Area";
	
	@Disabled("@BeforeAll")
	static void setupImagesAndTables() {
		
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
		
		OUTPUT_FILE = new File(System.getProperty("user.home") + "/Desktop/Test.xlsx");
		if (OUTPUT_FILE.exists()) {
			OUTPUT_FILE.delete();
		}
	}
	
			
	@Disabled
	void testSaveTableToWorkbookSheetResultsTableFileObjectBooleanBoolean() {
		
		
		
		ExcelFunctions.saveTableToWorkbookSheet(RESULTS_TABLE, OUTPUT_FILE, "", false);	
		
		File testFile = new File("/excel.macro.extensions/src/test/resources/xlsx_files/Test1.xlsx");
			
		ExcelFunctions.saveTableToWorkbookSheet(RESULTS_TABLE, OUTPUT_FILE, "SecondSheet", false);
//		
//		System.out.println(testFile.length());
//		assertFalse(OUTPUT_FILE == null);
//		try {
//			assertTrue(FileUtils.contentEquals(OUTPUT_FILE, testFile), "the 2 xlsx files should be equal");
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
		
		//OUTPUT_FILE.delete();
		
//		assertTrue(ExcelFunctions.saveTableToWorkbookSheet(RESULTS_TABLE, OUTPUT_FILE, 0, true, false));
//		assertTrue(ExcelFunctions.saveTableToWorkbookSheet(RESULTS_TABLE, OUTPUT_FILE, 1, true, true));
//		assertTrue(ExcelFunctions.saveTableToWorkbookSheet(SUMMARY_TABLE, OUTPUT_FILE, 4, false, false));
//		assertTrue(ExcelFunctions.saveTableToWorkbookSheet(SUMMARY_TABLE, OUTPUT_FILE, "FIRST", true, false));
//		assertTrue(ExcelFunctions.saveTableToWorkbookSheet(SUMMARY_TABLE, OUTPUT_FILE, "LAST", true, false));
//		assertTrue(ExcelFunctions.saveTableToWorkbookSheet(RESULTS_TABLE, OUTPUT_FILE, "", false, false));
		//OUTPUT_FILE.delete();
	}

	@Disabled
	void testSaveTableToWorkbookSheetResultsTableStringObjectBooleanBoolean() {
		fail("Not yet implemented");
	}

	@Disabled
	void testSaveTableToWorkbookSheetStringStringObjectBooleanBoolean() {
		fail("Not yet implemented");
	}

	@Disabled
	void testSaveAllOpenTablesToWorkbookFileBooleanBoolean() {
		fail("Not yet implemented");
	}

	@Disabled
	void testSaveAllOpenTablesToWorkbookStringBooleanBoolean() {
		fail("Not yet implemented");
	}

	@Disabled
	void testAddImageToWorkbookSheet() {
		fail("Not yet implemented");
	}

	@Disabled
	void testAppendArrayAsExcelRowNumberArrayStringObjectInt() {
		fail("Not yet implemented");
	}

	@Disabled
	void testAppendArrayAsExcelRowStringArrayStringObjectInt() {
		fail("Not yet implemented");
	}

	@Disabled
	void testAppendArrayAsExcelColumnNumberArrayStringObjectIntString() {
		fail("Not yet implemented");
	}

	@Disabled
	void testAppendArrayAsExcelColumnStringArrayStringObjectIntString() {
		fail("Not yet implemented");
	}

	@Disabled
	void testSetColumnDataFormat() {
		fail("Not yet implemented");
	}

}
