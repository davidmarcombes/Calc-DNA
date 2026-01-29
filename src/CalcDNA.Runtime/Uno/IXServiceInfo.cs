namespace CalcDNA.Runtime.Uno;

/// <summary>
/// Provides information about a UNO service implementation.
/// Mirrors com.sun.star.lang.XServiceInfo from UNO.
/// </summary>
public interface IXServiceInfo
{
    /// <summary>
    /// Returns the implementation name of this service.
    /// </summary>
    string getImplementationName();

    /// <summary>
    /// Checks if this service supports the specified service name.
    /// </summary>
    bool supportsService(string serviceName);

    /// <summary>
    /// Returns all service names this implementation supports.
    /// </summary>
    string[] getSupportedServiceNames();
}
