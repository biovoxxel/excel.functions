/**
 * 
 */
package excel.functions.plugins;

import java.io.File;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.scijava.command.Command;
import org.scijava.plugin.Parameter;
import org.scijava.plugin.Plugin;

import excel.functions.macro.BioVoxxelMacroExtensionDescriptor;
import excel.functions.utils.ExcelUtils;
import ij.macro.MacroExtension;
import ij.measure.ResultsTable;

/**
 * @author BioVoxxel
 *
 */
@Plugin(type = Command.class, menuPath = "Plugins>Excel Functions>Append table as rows")
public class AppendTableAsRows implements Command, BioVoxxelMacroExtensionDescriptor {

	   
	@Parameter(label = "Results table name", description = "if this field is empty the active table will be automatically used", required = false)
    private ResultsTable resultsTable;
    
	@Parameter(required = true)
	private File file;
	 
    @Parameter(label = "Sheet name", description = "if empty the sheet will receive the name of the table used truncated at 30 characters if the title is >30 char", required = false)
    private String sheetName;
      
    @Parameter(label = "Include table headings while appending")
    private Boolean includeTableHeadings = true;
    
       
    @Override
	public void run() {
		
		if (resultsTable == null) {
			resultsTable = ExcelUtils.getActiveTable();
		}
		
		if (sheetName.equals("")) {
			sheetName = resultsTable.getTitle();
		}
		
		appendTableAsRows(resultsTable, file, sheetName, includeTableHeadings);
	}
    
    
   protected static void appendTableAsRows(ResultsTable ijResultsTable, File workbookFile, String sheetNameOrIndexString, Boolean includeHeadings) {		
			
		String[][] table2DArray = ExcelUtils.getIJTableAsRowColumn2DArray(ijResultsTable, includeHeadings);
	
		Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		System.out.println("Workbook = " + workbook);
		
		if (workbook == null) {
			System.out.println("Aborting due to nonexisting workbook");
			return;	//stop processing if no workbook created or access blocked
		}
		
		Sheet sheet = ExcelUtils.getSheet(workbook, sheetNameOrIndexString);
		System.out.println("sheet = " + sheet.getSheetName());
		
		int rowCount = ExcelUtils.getRowCount(sheet);
		
		ExcelUtils.write2DArrayToSheet(sheet, rowCount, 0, table2DArray);
		
		ExcelUtils.saveWorkbook(workbook, workbookFile);
		
	}
	
	
	@Override
	public String runFromMacro(Object[] parameters) {
			
		ExcelUtils.logMacroParameters(parameters);
		
		ResultsTable resultsTable = ExcelUtils.getIJResultsTable(parameters[0].toString());
				
		File workbookFile = new File(ExcelUtils.fixFilePath(parameters[1].toString()));
		
		String sheetName = parameters[2].toString();
		if (sheetName.equals("")) {
			sheetName = resultsTable.getTitle();
		}
		
		boolean includeHeadings;
		if (parameters[3].toString().equals("1") || parameters[3].toString().equalsIgnoreCase("true")) {
			includeHeadings = true;
		} else {
			includeHeadings = false;
		}
		
		appendTableAsRows(resultsTable, workbookFile, sheetName, includeHeadings);
		
		return "";
	}

	@Override
	public int[] parameterTypes() {
		return new int[] { MacroExtension.ARG_STRING, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING };
	}

	@Override
	public String description() {
		
		return "Appends the specified or otherwise active table to an existing worksheet as new rows just below the existing rows";
	}

	@Override
	public String parameters() {
		return "ijResultsTableTitle, filePathToExcelWorkbook, sheetNameOrIndex, includeTableHeadings";
	}

}
