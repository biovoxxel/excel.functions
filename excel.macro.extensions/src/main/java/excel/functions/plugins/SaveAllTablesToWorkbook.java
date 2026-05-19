/**
 * 
 */
package excel.functions.plugins;

import java.io.File;

import org.scijava.command.Command;
import org.scijava.plugin.Parameter;
import org.scijava.plugin.Plugin;

import excel.functions.macro.BioVoxxelMacroExtensionDescriptor;
import excel.functions.utils.ExcelUtils;
import ij.WindowManager;
import ij.macro.MacroExtension;
import ij.measure.ResultsTable;

/**
 * @author BioVoxxel
 *
 */

@Plugin(type = Command.class, menuPath = "Plugins>Excel Functions>Save all tables to workbook")
public class SaveAllTablesToWorkbook implements Command, BioVoxxelMacroExtensionDescriptor {

	@Parameter
	File file;
		
	@Parameter(label = "Include table headings")
    private Boolean includeHeadings;
	
	
	@Override
	public void run() {
		
		saveAllOpenTablesAsWorkbookSheets(file, includeHeadings);
	}
	
	
	
	protected static void saveAllOpenTablesAsWorkbookSheets(File workbookFile, boolean includeHeadings) {
		
		String[] nonImageWindowNames = WindowManager.getNonImageTitles();
		System.out.println("Non image windows: " + nonImageWindowNames);
		
		for (int window = 0; window < nonImageWindowNames.length; window++) {
			
			ResultsTable currentTable = ResultsTable.getResultsTable(nonImageWindowNames[window]);
			System.out.println("Current results table = " + currentTable);
			
			if (currentTable != null) {
				SaveTableAsWorksheet.saveTableAsWorkbookSheet(currentTable, workbookFile, nonImageWindowNames[window], includeHeadings);
			}
		}		
	}
	
	
	@Override
	public String runFromMacro(Object[] parameters) {
		
		ExcelUtils.logMacroParameters(parameters);
		
		File workbookFile = new File(ExcelUtils.fixFilePath(parameters[0].toString()));
		
		boolean includeHeadings;
		if (parameters[1].toString().equals("1") || parameters[1].toString().equalsIgnoreCase("true")) {
			includeHeadings = true;
		} else {
			includeHeadings = false;
		}
		
		saveAllOpenTablesAsWorkbookSheets(workbookFile, includeHeadings);
		
		return "";
	}

	@Override
	public int[] parameterTypes() {
		 return new int[] { MacroExtension.ARG_STRING, MacroExtension.ARG_STRING };
	}

	@Override
	public String description() {
		return "Saves all open IJ results table as individual sheets to an XLSX Workbook.\n"
				+ "The name of the table is used as sheet name and trunkated after 31 characters";
	}

	@Override
	public String parameters() {
		return "filePathToExcelWorkbook, includeColumnHeadings";
	}

}
