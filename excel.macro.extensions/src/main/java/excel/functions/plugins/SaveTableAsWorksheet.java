/**
 * 
 */
package excel.functions.plugins;

import java.awt.Component;
import java.awt.Frame;
import java.awt.Window;
import java.io.File;
import java.util.logging.Logger;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

@Plugin(type = Command.class, menuPath = "Plugins>Excel Functions>Save table as worksheet")
public class SaveTableAsWorksheet implements Command, BioVoxxelMacroExtensionDescriptor {

	protected static final Logger logger = Logger.getLogger(ExcelUtils.class.getName());
    

	@Parameter(label = "Results table name", description = "if this field is empty the active table will be automatically used", required = false)
    private ResultsTable selectedTable = null;
    
	@Parameter(required = true)
	private File file;
	 
    @Parameter(label = "Sheet name", description = "if empty the sheet will receive the name of the table used truncated at 30 characters if the title is >30 char", required = false)
    private String sheetName;
   
    @Parameter(label = "Include column headings")
    private Boolean includeHeadings = true;
    
   
      
    @Override
	public void run() {
    	Frame activeTable = WindowManager.getFrontWindow();
    	
    	if (activeTable != null) {
    		   		
    		selectedTable = ResultsTable.getResultsTable(activeTable.getTitle());
    		
    		if (selectedTable != null) {
    			saveTableAsWorkbookSheet(selectedTable, file, sheetName, includeHeadings);							
			} else {
				JOptionPane.showMessageDialog(null, "Active Window is not a table", "No Table", JOptionPane.ERROR_MESSAGE);
			}
    		
		} else {
			JOptionPane.showMessageDialog(null, "No table window found", "No Table", JOptionPane.ERROR_MESSAGE);
		}
    		
	}
	
	
	protected static void saveTableAsWorkbookSheet(ResultsTable ijResultsTable, File workbookFile, String sheetName, Boolean includeColumnHeadings) {
		
		if (ijResultsTable == null) {
			JOptionPane.showMessageDialog(null, "No results table found", "No table found", JOptionPane.WARNING_MESSAGE);
			return;	//stop if an ImageJ table cannot be determined 
		}
		logger.finest("Current results table = " + ijResultsTable);
		
		String[][] table2DArray = ExcelUtils.getIJTableAsRowColumn2DArray(ijResultsTable, includeColumnHeadings);
			
		Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		logger.finest("Workbook = " + workbook);
		
		if (workbook == null) {
			logger.finest("Aborting due to nonexisting workbook");
			return;	//stop processing if no workbook created or access blocked
		}
		
		Sheet sheet = ExcelUtils.getUniqueSheet(workbook, sheetName);
		logger.finest("Current sheet = " + sheet);
		
		ExcelUtils.write2DArrayToSheet(sheet, 0, 0, table2DArray);
		
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
		
		saveTableAsWorkbookSheet(resultsTable, workbookFile, sheetName, includeHeadings);
		
		return "";
	}



	/**
     * We need to define a list of parameter types
     * @return int array with parameter types as defined in the MacroExtension class.
     */
    @Override
    public int[] parameterTypes() {
        return new int[] { MacroExtension.ARG_STRING, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING };
    }

    /**
     * We should define a description for users.
     * @return algorithm/parameter description
     */
    @Override
    public String description() {
        return "The active or otherwise specified table will be saved as a single sheet in an Excel workbook (.xlsx) file "
        		+ "under the specified name or the ImageJ table name otherwise.\n"
        		+ "This method will always create unique sheets in case the specified sheet name is already existing.\n"
        		+ "It will never overwrite any existing sheet.\n"
        		+ "If A table should be appended to an existing table use the corresponding methods.";
    }


	/**
     * We should provide a user readable list of parameters
     * @return list of parameters
     */
    @Override
    public String parameters() {
        return "ijResultsTableTitle, filePathToExcelWorkbook, sheetName, includeColumnHeadings";
    }
	
	
}
