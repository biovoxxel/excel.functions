package excel.functions.plugins;

import excel.functions.macro.BioVoxxelMacroExtensionDescriptor;

/**
 * 
 * @author BioVoxxel
 *
 */

public interface BioVoxxelPlugins {

    // macro extensions should have a prefix in order to prevent
    // conflicts between different extensions
    String macroExtensionPrefix = "xlsx_";

    // list of all available plugins
    BioVoxxelMacroExtensionDescriptor[] list = {
    		new AddImageToWorksheet(),
            new SaveTableAsWorksheet(),
            new SaveAllTablesToWorkbook(),
            new ImportXlsxWorksheetAsTable(), 
            new DeleteWorkbookSheet(),
            new SetColumnDataFormat(),
            new AppendTableAsColumns(),
            new AppendTableAsRows(),
            new GetAllWorkbookSheetNames(),
            new AppendArrayAsExcelRow(),
            new AppendArrayAsExcelColumn(),
            new AddExcelChartFromPlot()
    };
	
}
