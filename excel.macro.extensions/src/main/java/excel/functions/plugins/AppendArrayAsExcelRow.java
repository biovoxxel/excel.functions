/**
 * 
 */
package excel.functions.plugins;

import java.io.File;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import excel.functions.macro.BioVoxxelMacroExtensionDescriptor;
import excel.functions.utils.ExcelUtils;
import ij.macro.MacroExtension;

/**
 * @author BioVoxxel
 * 
 */

public class AppendArrayAsExcelRow implements BioVoxxelMacroExtensionDescriptor {
	
	protected static void appendArrayAsTableRow(String[] array, File workbookFile, String sheetNameOrIndexString, int startingColumn) {
		
		String[][] table2DArray = ExcelUtils.convertInRowArray(array);
	
		Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		System.out.println("Workbook = " + workbook);
		
		if (workbook == null) {
			System.out.println("Aborting due to nonexisting workbook");
			return;	//stop processing if no workbook created or access blocked
		}
		
		Sheet sheet = ExcelUtils.getSheet(workbook, sheetNameOrIndexString);
		
		int rowCount = ExcelUtils.getRowCount(sheet);
		
		ExcelUtils.write2DArrayToSheet(sheet, rowCount, startingColumn, table2DArray);
		
		ExcelUtils.saveWorkbook(workbook, workbookFile);	
	}

	
	
	@Override
	public String runFromMacro(Object[] parameters) {
		
		ExcelUtils.logMacroParameters(parameters);
		
		Object[] inputDataArray = (Object[])parameters[0];
		String[] inputDataArrayAsStringArray = new String[inputDataArray.length];
		
		for (int dataPoint = 0; dataPoint < inputDataArray.length; dataPoint++) {
			inputDataArrayAsStringArray[dataPoint] = inputDataArray[dataPoint].toString();
		}
		
		File workbookFile = new File(ExcelUtils.fixFilePath(parameters[1].toString()));
		
		String sheetName = parameters[2].toString();
		
		int startingRow = Integer.parseInt(parameters[3].toString());
				
		appendArrayAsTableRow(inputDataArrayAsStringArray, workbookFile, sheetName, startingRow);
		
		return "";
	}

	@Override
	public int[] parameterTypes() {
		return new int[] { MacroExtension.ARG_ARRAY, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING };
	}

	@Override
	public String description() {
		
		return "Appends data from an array as ROW after all existing rows starting at a defined COLUMN.";
	}

	@Override
	public String parameters() {
		return "array, filePathToExcelWorkbook, sheetNameOrIndex, startingColumn";
	}

}
