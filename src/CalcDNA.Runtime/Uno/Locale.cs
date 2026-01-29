namespace CalcDNA.Runtime.Uno;

/// <summary>
/// Represents a locale used by LibreOffice for internationalization.
/// Mirrors com.sun.star.lang.Locale from UNO.
/// </summary>
public struct Locale
{
    /// <summary>
    /// ISO 639 language code (e.g., "en", "de", "fr").
    /// </summary>
    public string Language { get; set; }

    /// <summary>
    /// ISO 3166 country code (e.g., "US", "DE", "FR").
    /// </summary>
    public string Country { get; set; }

    /// <summary>
    /// Variant code for special cases.
    /// </summary>
    public string Variant { get; set; }

    public Locale(string language, string country, string variant)
    {
        Language = language ?? "";
        Country = country ?? "";
        Variant = variant ?? "";
    }

    public static Locale Default => new("", "", "");
}
