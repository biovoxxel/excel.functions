/**
 * 
 */
package excel.functions.macro;

/**
 * @author BioVoxxel
 *
 */
public interface BioVoxxelMacroExtensionDescriptor {
	String runFromMacro(Object[] parameters);
    int[] parameterTypes();
    String description();
    String parameters();
}
