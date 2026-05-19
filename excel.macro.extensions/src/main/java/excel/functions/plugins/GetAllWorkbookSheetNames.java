/**
 * 
 */
package excel.functions.plugins;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

import excel.functions.macro.BioVoxxelMacroExtensionDescriptor;
import excel.functions.utils.ExcelUtils;
import ij.macro.MacroExtension;

/**
 * @author BioVoxxel
 *
 */

public class GetAllWorkbookSheetNames implements BioVoxxelMacroExtensionDescriptor {

	protected static String getAllWorkbookSheetNames(File workbookFile) throws IOException {
		
		Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		System.out.println("Workbook = " + workbook);
		
		if (workbook == null) {
			System.out.println("Aborting due to nonexisting workbook");
			return "";	//stop processing if no workbook created or access blocked
		}
		
		List<String> workbookSheetNames = ExcelUtils.getAllSheetNames(workbook);
		
		String sheetNames = "";
		
		for (int sheetIndex = 0; sheetIndex < workbookSheetNames.size(); sheetIndex++) {
			if (sheetIndex == 0) {
				sheetNames = workbookSheetNames.get(sheetIndex);				
			} else {
				sheetNames = sheetNames + "," + workbookSheetNames.get(sheetIndex);	
			}
		}
		workbook.close();
		
		return sheetNames;
	}
	
	
	@Override
	public String runFromMacro(Object[] parameters) {
		File workbookFile = new File(ExcelUtils.fixFilePath(parameters[0].toString()));
		
		try {
			return getAllWorkbookSheetNames(workbookFile);
		} catch (IOException e) {
			e.printStackTrace();
			return "";
		}
		
	}

	@Override
	public int[] parameterTypes() {
		return new int[] { MacroExtension.ARG_STRING };
	}

	@Override
	public String description() {
		
		return "Returns the names of all existing sheets in a specified workbook as concatenated String separated by commas.\n"
				+ "Use 'split(string, delimiter);' to split sheet names into a String[] array and individually access the names.";
	}

	@Override
	public String parameters() {
		
		return "filePathToXlsxFile";
	} 

}
