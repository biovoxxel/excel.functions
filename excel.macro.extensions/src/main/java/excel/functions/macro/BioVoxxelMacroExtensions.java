/**
 * 
 */
package excel.functions.macro;

import java.util.logging.Level;

import org.scijava.command.Command;
import org.scijava.plugin.Parameter;
import org.scijava.plugin.Plugin;

import excel.functions.plugins.BioVoxxelPlugins;
import excel.functions.utils.ExcelUtils;
import ij.macro.ExtensionDescriptor;
import ij.macro.Functions;
import ij.macro.MacroExtension;

/**
 * @author BioVoxxel
 *
 */

@Plugin(type = Command.class, menuPath = "Plugins>Excel Functions>Excel Macro Extensions")
public class BioVoxxelMacroExtensions implements Command, MacroExtension {
	
	@Parameter(label = "Debug logging", required = false)
	private boolean debuglogging;  
	
	@Override
	public void run() {
		// Activate this class as handler for macro extensions
		if (debuglogging) {
			ExcelUtils.setupLogger(Level.ALL);			
		} else {
			ExcelUtils.setupLogger(Level.OFF);		
		}
		Functions.registerExtensions(this);
	}

	@Override
	public String handleExtension(String name, Object[] args) {
		BioVoxxelMacroExtensionDescriptor[] pluginList = BioVoxxelPlugins.list;

		String readout = "";
        // go through plugin list and check if we know the called plugin
        for (BioVoxxelMacroExtensionDescriptor plugin : pluginList) {
            String command = BioVoxxelPlugins.macroExtensionPrefix + plugin.getClass().getSimpleName();

            // if name matches exactly
            if (command.compareTo(name) == 0) {
                // call the plugin
                readout = plugin.runFromMacro(args);
                break;
            }
        }
		return readout;
	}

	@Override
	public ExtensionDescriptor[] getExtensionFunctions() {
		
		BioVoxxelMacroExtensionDescriptor[] pluginList = BioVoxxelPlugins.list;

        ExtensionDescriptor[] result = new ExtensionDescriptor[pluginList.length];

        int i = 0;
        // formulate a list of ExtensionDescriptors describing all command this class can handle
        for (BioVoxxelMacroExtensionDescriptor plugin : pluginList) {
            String call = BioVoxxelPlugins.macroExtensionPrefix + plugin.getClass().getSimpleName();
            result[i] = new ExtensionDescriptor(call, plugin.parameterTypes(), this);
            i++;
        }

        // hand over the list to ImageJs macro interpreter
        return result;
	}

}
