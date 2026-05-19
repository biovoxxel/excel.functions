/**
 * 
 */
package excel.functions.plugins;

import java.io.File;
import java.io.IOException;

import javax.swing.JOptionPane;

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

@Plugin(type = Command.class, menuPath = "Plugins>Excel Functions>Import worksheet as table")
public class ImportXlsxWorksheetAsTable implements Command, BioVoxxelMacroExtensionDescriptor {

	@Parameter(required = true)
    private File file;
	
	@Parameter(label = "Workbook sheet name or index", description = "", required = true)
    private String sheetName;
	
	@Parameter(label = "Use 1st row as headings", description = "", required = false)
    private Boolean useFirstRowAsHeadings;

	@Override
	public void run() {
		
		if (!file.exists()) {
			JOptionPane.showMessageDialog(null, "Specified file is not existing!", "File not found", JOptionPane.WARNING_MESSAGE);
			return;
		}
				
		importXlsxWorksheetAsTable(file, sheetName, useFirstRowAsHeadings);
		
	}

	
	protected static void importXlsxWorksheetAsTable(File workbookFile, String sheetNameOrIndexString, boolean useFirstRowAsHeading) {
		
		Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		System.out.println("Workbook = " + workbook);
		
		if (workbook == null) {
			System.out.println("Aborting due to nonexisting workbook");
			return;	//stop processing if no workbook created or access blocked
		}
		
		
		Sheet sheet = ExcelUtils.getSheet(workbook, sheetNameOrIndexString);
				
		if (sheet == null) {
			JOptionPane.showMessageDialog(null, "Sheet could not be found", "No such sheet", JOptionPane.WARNING_MESSAGE);
			return;
		}
		
		if (ExcelUtils.getRowCount(sheet) > 0 && ExcelUtils.getColumnCount(sheet) > 0) {
			String[][] table2DArray = ExcelUtils.getSheetAsRowColumn2DArray(sheet);

			ResultsTable resultsTable = ExcelUtils.createResultsTableFrom2DArray(table2DArray, useFirstRowAsHeading);
			
			resultsTable.show(sheet.getSheetName());			
		}
		
		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	@Override
	public String runFromMacro(Object[] parameters) {
				
		ExcelUtils.logMacroParameters(parameters);
		
		File workbookFile = new File(ExcelUtils.fixFilePath(parameters[0].toString()));
		
		String sheetName = parameters[1].toString();
		
		boolean useFirstRowAsHeading;
		if (parameters[2].toString().equals("1") || parameters[2].toString().equalsIgnoreCase("true")) {
			useFirstRowAsHeading = true;
		} else {
			useFirstRowAsHeading = false;
		}
		
		importXlsxWorksheetAsTable(workbookFile, sheetName, useFirstRowAsHeading);
		
		return "";
	}


	

	@Override
	public int[] parameterTypes() {
		return new int[] { MacroExtension.ARG_STRING, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING };
	}

	@Override
	public String description() {
		return "Loads a specified spreadsheet from an Excel workbook."
				+ "Sheetnames can be specified as the real names or as sheet indices (0-based).";
	}

	@Override
	public String parameters() {
		 return "filePathToExcelWorkbook, sheetNameOrIndex, useFirstRowAsHeasing";
	}
	
	
}
