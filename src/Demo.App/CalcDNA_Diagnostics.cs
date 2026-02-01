namespace Demo.App;

internal static class CalcDNA_Diagnostics
{
    [System.Runtime.CompilerServices.ModuleInitializer]
    public static void OnAssemblyLoad()
    {
        try
        {
            System.IO.File.AppendAllText("/tmp/calcdna_debug.log",
                $"[{DateTime.UtcNow:yyyy-MM-dd HH:mm:ss.fff}] Demo.App assembly loaded, PID={System.Diagnostics.Process.GetCurrentProcess().Id}\n");
        }
        catch { }
    }
}
