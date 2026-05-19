/**
 * 
 */
package excel.functions.plugins;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import excel.functions.macro.BioVoxxelMacroExtensionDescriptor;
import excel.functions.utils.ExcelUtils;
import ij.macro.MacroExtension;

/**
 * @author BioVoxxel
 *
 */
public class AppendArrayAsExcelColumn implements BioVoxxelMacroExtensionDescriptor {


	
	protected static void appendArrayAsTableColumn(String[] array, File workbookFile, String sheetNameOrIndexString, int startingRow) {
		
		String[][] table2DArray = ExcelUtils.convertInColumnArray(array);
	
		Workbook workbook;
		try {
			workbook = WorkbookFactory.create(workbookFile);
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
			return;
		} catch (IOException e) {
			e.printStackTrace();
			return;
		}
		System.out.println("Workbook = " + workbook);
		
		Sheet sheet = ExcelUtils.getSheet(workbook, sheetNameOrIndexString);
		
		int columnCount = ExcelUtils.getColumnCount(sheet);
		
		ExcelUtils.write2DArrayToSheet(sheet, startingRow, columnCount, table2DArray);
		
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
				
		appendArrayAsTableColumn(inputDataArrayAsStringArray, workbookFile, sheetName, startingRow);
		
		return "";
	}

	@Override
	public int[] parameterTypes() {
		return new int[] { MacroExtension.ARG_ARRAY, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING };
	}

	@Override
	public String description() {
		
		return "Appends data from an array as COLUMN after all existing columns starting at a defined ROW.";
	}

	@Override
	public String parameters() {
		return "array, filePathToExcelWorkbook, sheetNameOrIndex, startingRow";
	}

}
