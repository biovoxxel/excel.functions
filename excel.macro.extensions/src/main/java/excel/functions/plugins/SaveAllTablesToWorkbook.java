/**
 * 
 */
package excel.functions.plugins;

import java.io.File;
import java.util.logging.Logger;

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

	protected static final Logger logger = Logger.getLogger(ExcelUtils.class.getName());
	
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
		logger.finest("Non image windows: " + nonImageWindowNames);
		
		for (int window = 0; window < nonImageWindowNames.length; window++) {
			
			ResultsTable currentTable = ResultsTable.getResultsTable(nonImageWindowNames[window]);
			logger.finest("Current results table = " + currentTable);
			
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
