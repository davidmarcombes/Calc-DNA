namespace CalcDNA.Runtime.Uno;

/// <summary>
/// Interface for Calc Add-In functions providing display names and descriptions.
/// Mirrors com.sun.star.sheet.XAddIn from UNO.
/// </summary>
public interface IXAddIn
{
    /// <summary>
    /// Returns the internal programmatic name of a function.
    /// </summary>
    string getProgrammaticFuntionName(string displayName);

    /// <summary>
    /// Returns the user-visible display name of a function.
    /// </summary>
    string getDisplayFunctionName(string programmaticName);

    /// <summary>
    /// Returns the description of a function.
    /// </summary>
    string getFunctionDescription(string programmaticName);

    /// <summary>
    /// Returns the display name of a parameter.
    /// </summary>
    string getDisplayArgumentName(string programmaticFunctionName, int argumentIndex);

    /// <summary>
    /// Returns the description of a parameter.
    /// </summary>
    string getArgumentDescription(string programmaticFunctionName, int argumentIndex);

    /// <summary>
    /// Returns the category name for a function.
    /// </summary>
    string getProgrammaticCategoryName(string programmaticFunctionName);

    /// <summary>
    /// Returns the display category name for a function.
    /// </summary>
    string getDisplayCategoryName(string programmaticFunctionName);
}
