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

@Plugin(type = Command.class, menuPath = "Plugins>Excel Functions>Delete worksheet")
public class DeleteWorkbookSheet implements Command, BioVoxxelMacroExtensionDescriptor {
	
	
	@Parameter(required = true)
    private File file;
	
	@Parameter(label = "Sheet name or index", initializer = "", style = "", required = true)
	private String sheetName;
	

	
	@Override
	public void run() {
		
		deleteWorkbookSheet(file, sheetName);
		
	}

	protected static void deleteWorkbookSheet(File workbookFile, String sheetNameOrIndexString) {
			
		Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		System.out.println("Workbook = " + workbook);
		
		if (workbook == null) {
			System.out.println("Aborting due to nonexisting workbook");
			return;	//stop processing if no workbook created or access blocked
		}
		
		Sheet sheet = ExcelUtils.getSheet(workbook, sheetNameOrIndexString);
		
		int sheetIndex = workbook.getSheetIndex(sheet);
				
		if (sheetIndex >= 0) {
			workbook.removeSheetAt(sheetIndex);
		}
												
		ExcelUtils.saveWorkbook(workbook, workbookFile);
		
	}
	
	
	@Override
	public String runFromMacro(Object[] parameters) {
				
		ExcelUtils.logMacroParameters(parameters);
		
		File workbookFile = new File(ExcelUtils.fixFilePath(parameters[0].toString()));
		
		String sheetName = parameters[1].toString();
		
		deleteWorkbookSheet(workbookFile, sheetName);
		
		return "";
	}

	@Override
	public int[] parameterTypes() {
		 return new int[] { MacroExtension.ARG_STRING, MacroExtension.ARG_STRING };
	}

	@Override
	public String description() {
		return "Deletes a specified sheet in a specified Excel workbook. The sheet can be given by name or 0-based sheet index";
	}

	@Override
	public String parameters() {
		
		return "filePathToExcelWorkbook, sheetNameOrIndex";
	}

}
