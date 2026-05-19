/**
 * 
 */
package excel.functions.plugins;

import java.io.File;
import java.io.IOException;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.scijava.command.Command;
import org.scijava.plugin.Parameter;
import org.scijava.plugin.Plugin;

import excel.functions.macro.BioVoxxelMacroExtensionDescriptor;
import excel.functions.utils.ExcelUtils;
import ij.ImagePlus;
import ij.WindowManager;
import ij.macro.MacroExtension;

/**
 * @author BioVoxxel
 *
 */

@Plugin(type = Command.class, menuPath = "Plugins>Excel Functions>Add image to worksheet")
public class AddImageToWorksheet implements Command, BioVoxxelMacroExtensionDescriptor {
	
	@Parameter
	ImagePlus imagePlus;
	
	@Parameter
    File file;
	
	@Parameter (label = "Sheet name or index", description = "specify the Excel worksheet name or the 0-based index of the worksheet,"
			+ "non-existing sheets will be automatically created")
	String sheetName;
	
	@Parameter (description = "0-based Excel table column will be the upper left corner of the image added")
	int column;
	
	@Parameter (description = "0-based Excel table row will be the upper left corner of the image added")
	int row;
	
	@Override
	public void run() {
	
		addImageToWorksheet(imagePlus, file, sheetName, column, row);
	}
	
	protected static void addImageToWorksheet(ImagePlus imagePlus, File workbookFile, String sheetNameOrIndexString, int column, int row) {
				
		System.out.println("ImagePlus = " + imagePlus);
		
		if (imagePlus == null) {
			JOptionPane.showMessageDialog(null, "Specified image could not be found\n"
					+ "or there is no image open", "No such image", JOptionPane.WARNING_MESSAGE);
			return;
		}
	
		
		byte[] imageData = null;
		try {
			imageData = ExcelUtils.getImageAsByteArray(imagePlus);
		} catch (IOException e) {
			e.printStackTrace();
		}
				

		Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		System.out.println("Workbook = " + workbook);
		
		if (workbook == null) {
			System.out.println("Aborting due to nonexisting workbook");
			return;	//stop processing if no workbook created or access blocked
		}
		
		Sheet sheet = ExcelUtils.getSheet(workbook, sheetNameOrIndexString);	
		System.out.println("Sheet = " + sheet);
		
		if (imageData != null) {
			int pictureIndex = workbook.addPicture(imageData, Workbook.PICTURE_TYPE_PNG);
			
			CreationHelper helper = workbook.getCreationHelper();
			
			//add a picture shape
			ClientAnchor anchor = helper.createClientAnchor();
			
			//specify the column and row in which the image will be saved. -1 as an input leads to automatic location
			int columnPictureLocation = column;
			Row firstRow = sheet.getRow(0);
			if (firstRow == null) {
				sheet.createRow(0);
			}
			
			if(columnPictureLocation < 0) {	
				columnPictureLocation = sheet.getRow(0).getLastCellNum();			
			}
			
			int rowPictureLocation = row;
			if (rowPictureLocation < 0) {
				rowPictureLocation = sheet.getLastRowNum()+1;
			}
			
			//set top-left corner of the picture,
			//subsequent call of Picture#resize() will operate relative to it
			anchor.setCol1(columnPictureLocation);
			anchor.setRow1(rowPictureLocation);
			
			// Create the drawing patriarch.  This is the top level container for all shapes.
			Drawing<?> drawing =  sheet.createDrawingPatriarch();
			
			Picture pict = drawing.createPicture(anchor, pictureIndex);
			
			//auto-size picture relative to its top-left corner
			pict.resize();
			
			ExcelUtils.saveWorkbook(workbook, workbookFile);
		}
		
	}

	
	

	@Override
	public String runFromMacro(Object[] parameters) {
	
		ExcelUtils.logMacroParameters(parameters);
		
		ImagePlus image = WindowManager.getImage((int)Math.round((Double)parameters[0]));
		File workbookFile = new File(ExcelUtils.fixFilePath(parameters[1].toString()));
		String sheetName = parameters[2].toString();
		if (sheetName.equals("")) {
			sheetName = image.getTitle();
		}
		int column = ((Double)parameters[3]).intValue();
		int row = ((Double)parameters[4]).intValue();
		
		addImageToWorksheet(image, workbookFile, sheetName, column, row);
		
		return "";
	}

	@Override
	public int[] parameterTypes() {
		 return new int[] { MacroExtension.ARG_NUMBER, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING, MacroExtension.ARG_NUMBER, MacroExtension.ARG_NUMBER };
	}

	@Override
	public String description() {
		
		return "Adds a specified image (using the negative image ID) or the active image if no image is given to the specified Excel worksheet "
				+ "in a specified workbook file or creates a new workbook file if the specified one is not existing.\n"
				+ "If a column and row is specified, the upper left corner of the image is at that position.\n"
				+ "If value for column = 0, the image will be placed after the last column with data.\n"
				+ "If row = 0, the image will be placed after the third row.";
	}

	@Override
	public String parameters() {
		 return "imageID, filePathToExcelWorkbook, sheetNameOrIndex, column, row";
	}

}
