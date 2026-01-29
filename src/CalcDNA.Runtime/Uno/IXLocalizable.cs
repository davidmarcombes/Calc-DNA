namespace CalcDNA.Runtime.Uno;

/// <summary>
/// Provides locale information for internationalization.
/// Mirrors com.sun.star.lang.XLocalizable from UNO.
/// </summary>
public interface IXLocalizable
{
    /// <summary>
    /// Sets the locale to be used by this object.
    /// </summary>
    void setLocale(Locale locale);

    /// <summary>
    /// Returns the currently set locale.
    /// </summary>
    Locale getLocale();
}
