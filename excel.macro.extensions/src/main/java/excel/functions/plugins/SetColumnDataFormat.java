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

/**
 * @author BioVoxxel
 *
 */
@Plugin(type = Command.class, menuPath = "Plugins>Excel Functions>Set column data format")
public class SetColumnDataFormat implements Command, BioVoxxelMacroExtensionDescriptor {

	protected static boolean CLOSE_WORKBOOK_AFTER_SAVING = true;
	
	@Parameter(required = false)
	private File file;
	 
    @Parameter(label = "Sheet name or index", description = "if empty the sheet will receive the name of the table used truncated at 30 characters if the title is >30 char", required = false)
    private String sheetName;
    
    @Parameter(label="Column index (0-based)")
    private int columnIndex;
    
    @Parameter(label="Data format", choices = {"General", "0", "0.00", "0%", "0.00%", "h:mm:ss", "0.00E+00"})
    private String dataFormat;
	

    @Override
	public void run() {
		
		setColumnDataFormat(file, sheetName, columnIndex, dataFormat);
		
	}
	
   protected static void setColumnDataFormat(File workbookFile, String sheetNameOrIndexString, int columnNumber, String dataFormatString) {
	   
	   Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		System.out.println("Workbook = " + workbook);
		
		if (workbook == null) {
			System.out.println("Aborting due to nonexisting workbook");
			return;	//stop processing if no workbook created or access blocked
		}
    	
    	Sheet sheet = ExcelUtils.getSheet(workbook, sheetNameOrIndexString);
    	System.out.println("Current sheet = " + sheet);
    	
    	ExcelUtils.setColumnDataFormat(sheet, columnNumber, dataFormatString);
    	
    	ExcelUtils.saveWorkbook(workbook, workbookFile);
    	
	}

	
	
	@Override
	public String runFromMacro(Object[] parameters) {
		ExcelUtils.logMacroParameters(parameters);
		
		File workbookFile = new File(ExcelUtils.fixFilePath(parameters[0].toString()));
		
		String sheetName = parameters[1].toString();
		
		int columnNumber = ((Double)parameters[2]).intValue();
		
		String dataFormatString = parameters[3].toString();
				
		setColumnDataFormat(workbookFile, sheetName, columnNumber, dataFormatString);
		
		return "";
	}

	
	@Override
	public int[] parameterTypes() {
		 return new int[] { MacroExtension.ARG_STRING, MacroExtension.ARG_STRING, MacroExtension.ARG_NUMBER, MacroExtension.ARG_STRING };
	}

	@Override
	public String description() {
		
		
		return "The column data format can de set for a defined workbook and sheet\n"
				+ "Possible formats are described under\n"
				+ "<a href=https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BuiltinFormats.html>BuiltinFormats</a>";
	}

	@Override
	public String parameters() {
		
		return "filePathToExcelWorkbook, sheetName, columnNumber, dataFormatString";
	}

	
	
}
