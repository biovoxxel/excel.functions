/**
 * 
 */
package excel.functions.plugins;

import java.awt.Window;
import java.io.File;
import java.util.logging.Logger;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.scijava.command.Command;
import org.scijava.plugin.Parameter;
import org.scijava.plugin.Plugin;

import excel.functions.macro.BioVoxxelMacroExtensionDescriptor;
import excel.functions.utils.ExcelUtils;
import ij.WindowManager;
import ij.gui.Plot;
import ij.gui.PlotWindow;
import ij.macro.MacroExtension;

/**
 * @author BioVoxxel
 * 
 * status: still experimental, since not all functions in Apache POI are fully functional
 *
 */
//not yet fully functional
@Plugin(type = Command.class, menuPath = "Plugins>Excel Functions>Add Excel chart from Plot (experimental)")
public class AddExcelChartFromPlot implements Command, BioVoxxelMacroExtensionDescriptor {

	protected static final Logger logger = Logger.getLogger(ExcelUtils.class.getName());
    
	@Parameter(required = true)
	private File file;
	 
    @Parameter(label = "Sheet name", description = "if empty the sheet will receive the name of the plot used truncated at 30 characters if the title is >30 char", required = false)
    private String sheetName;
    
    @Parameter(choices = {"AREA", "AREA3D", "LINE", "LINE3D", "SCATTER"})
    private String charttype;
    
    @Override
	public void run() {
    	
    	Plot plot = getActivePlot();
    	
		addExcelChartFromPlot(file, sheetName, plot, ChartTypes.valueOf(charttype));
		
	}

	private Plot getActivePlot() {
		Window window = WindowManager.getActiveWindow();
    	Plot plot = null;
    	if (window instanceof PlotWindow) {
    		plot = ((PlotWindow) window).getPlot();
		}
    	
    	if (plot == null) {
			JOptionPane.showMessageDialog(window, "Active Window is not a Plot", "Not a Plot", JOptionPane.WARNING_MESSAGE);
		}
		return plot;
	}
    
	protected static void addExcelChartFromPlot(File workbookFile, String sheetNameOrIndexString, Plot plot, ChartTypes chartType) {
		
		Workbook workbook = ExcelUtils.getWorkbook(workbookFile);
		logger.info("Workbook = " + workbook);
		
		if (workbook == null) {
			logger.info("Aborting due to nonexisting workbook");
			return;	//stop processing if no workbook created or access blocked
		}
		
		Sheet sheet = ExcelUtils.getSheet(workbook, sheetNameOrIndexString);
		logger.info("Current sheet = " + sheet);
		
		
		ExcelUtils.addExcelChartFromPlot(sheet, plot, chartType);
		
		ExcelUtils.saveWorkbook(workbook, workbookFile);
		
	}
	
	

	@Override
	public String runFromMacro(Object[] parameters) {
		
		ExcelUtils.logMacroParameters(parameters);
		
		File workbookFile = new File(ExcelUtils.fixFilePath(parameters[0].toString()));
				
		Plot plot = getActivePlot();
		String sheetName = parameters[1].toString();
		if (sheetName.equals("")) {
			sheetName = plot.getTitle();
		}
		
		ChartTypes chartType = ChartTypes.valueOf(parameters[2].toString());
		
		addExcelChartFromPlot(workbookFile, sheetName, plot, chartType);
		
		return "";
		
	}

	@Override
	public int[] parameterTypes() {
		return new int[] { MacroExtension.ARG_STRING, MacroExtension.ARG_STRING, MacroExtension.ARG_STRING };
	}

	@Override
	public String description() {
		return "This function takes the active plot window in ImageJ and converts the plot into an Excel chart."
				+ "Currently, the functionality and chart types are still limited but once saved in Excel,"
				+ "it is easy to change the chart type. More functionalities are to come."
				+ "Usable chart types need to be specified capitaliced and available are currently:"
				+ "AREA, AREA3D, LINE, LINE3D, SCATTER";
	}

	@Override
	public String parameters() {
		return "filePathToExcelWorkbook, sheetNameOrIndex, chartType";
	}

}
