package excel.functions.macro;

import java.util.ArrayList;
import java.util.List;

import org.fife.ui.autocomplete.BasicCompletion;
import org.fife.ui.autocomplete.CompletionProvider;
import org.scijava.plugin.Plugin;

import excel.functions.plugins.BioVoxxelPlugins;
import net.imagej.legacy.plugin.MacroExtensionAutoCompletionPlugin;

@Plugin(type = MacroExtensionAutoCompletionPlugin.class)
public class BioVoxxelMacroAutoCompletionExtension implements MacroExtensionAutoCompletionPlugin {

	@Override
	public List<BasicCompletion> getCompletions(CompletionProvider completionProvider) {
		
		ArrayList<BasicCompletion> completions = new ArrayList<BasicCompletion>();

		BioVoxxelMacroExtensionDescriptor[] pluginList = BioVoxxelPlugins.list;
		
		// go through plugins and provide an auto-complete entry for each
        for (BioVoxxelMacroExtensionDescriptor plugin : pluginList) {
            String commandName = "Ext." + BioVoxxelPlugins.macroExtensionPrefix + plugin.getClass().getSimpleName() + "(" + plugin.parameters() + ");";
            String description = "<b>" + commandName + "</b><br>" + plugin.description();
            completions.add(new BasicCompletion(completionProvider, commandName, null, description));
        }

        return completions;
		
	}
	
}
