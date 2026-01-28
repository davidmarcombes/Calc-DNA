using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace CalcDNA.CLI;

/// <summary>
/// Generates RDB (Registry Database) files from IDL using LibreOffice SDK tools.
/// RDB files contain compiled type information used by the UNO runtime.
/// </summary>
internal static class RdbGenerator
{
    /// <summary>
    /// Generates an RDB file from an IDL file using the LibreOffice SDK's unoidl-write tool.
    /// </summary>
    /// <param name="idlFilePath">Path to the input IDL file.</param>
    /// <param name="rdbOutputPath">Path for the output RDB file.</param>
    /// <param name="sdkPath">Optional explicit path to LibreOffice SDK. If null, will attempt to auto-detect.</param>
    /// <param name="logger">Logger for output messages.</param>
    public static void WriteRdb(string idlFilePath, string rdbOutputPath, string? sdkPath, Logger logger)
    {
        string? sdkActualPath = sdkPath ?? GuessLibreOfficeSdkPath(logger);

        if (string.IsNullOrEmpty(sdkActualPath) || !Directory.Exists(sdkActualPath))
        {
            throw new InvalidOperationException(
                "LibreOffice SDK path could not be determined. " +
                "Please install the LibreOffice SDK or specify the path explicitly using --sdk-path.");
        }

        logger.Debug($"Using LibreOffice SDK at: {sdkActualPath}", verbose: true);

        // Determine the correct executable name based on platform
        string exeName = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? "unoidl-write.exe"
            : "unoidl-write";

        var unoidlWrite = Path.Combine(sdkActualPath, "bin", exeName);

        if (!File.Exists(unoidlWrite))
        {
            throw new FileNotFoundException(
                $"unoidl-write not found in the LibreOffice SDK bin directory. Expected at: {unoidlWrite}",
                unoidlWrite);
        }

        // Locate required type definition files
        var typesRdb = FindTypesRdb(sdkActualPath);

        if (typesRdb == null)
        {
            throw new FileNotFoundException(
                "Could not find types.rdb or offapi.rdb in the LibreOffice SDK. " +
                "Ensure the SDK is properly installed.");
        }

        logger.Debug($"Using types RDB: {typesRdb}", verbose: true);

        // Build arguments for unoidl-write
        // Format: unoidl-write <registry>... <output.rdb>
        var arguments = $"\"{typesRdb}\" \"{idlFilePath}\" \"{rdbOutputPath}\"";

        logger.Debug($"Running: {unoidlWrite} {arguments}", verbose: true);

        // Run unoidl-write with proper process handling
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = unoidlWrite,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };

        // Ensure LibreOffice program directory (contains DLLs/shared libs) is available
        // so that unoidl-write can load required native libraries. We prepend the
        // program directory to PATH on Windows/Linux/macOS and also set the
        // platform-specific library path variables on Unix-like systems.
        var programDirCandidates = new[]
        {
            Path.GetFullPath(Path.Combine(sdkActualPath, "..", "program")),
            Path.Combine(sdkActualPath, "program"),
            Path.Combine(sdkActualPath, "..", "program", ".."),
            sdkActualPath
        };

        string? programDir = null;
        foreach (var cand in programDirCandidates)
        {
            try
            {
                if (!string.IsNullOrEmpty(cand) && Directory.Exists(cand))
                {
                    programDir = Path.GetFullPath(cand);
                    break;
                }
            }
            catch
            {
                // ignore invalid candidate
            }
        }

        if (programDir != null)
        {
            // Prepend to PATH for the child process
            var pathKey = "PATH";
            var existingPath = Environment.GetEnvironmentVariable(pathKey) ?? string.Empty;
            try
            {
                process.StartInfo.EnvironmentVariables[pathKey] = programDir + Path.PathSeparator + existingPath;
            }
            catch
            {
                // If EnvironmentVariables is not available for some reason, ignore
            }

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                logger.Debug($"Prepended LibreOffice program directory to PATH: {programDir}", verbose: true);
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                // macOS may require DYLD_LIBRARY_PATH for dynamic libraries
                var dyldKey = "DYLD_LIBRARY_PATH";
                var existingDyld = Environment.GetEnvironmentVariable(dyldKey) ?? string.Empty;
                try { process.StartInfo.EnvironmentVariables[dyldKey] = programDir + Path.PathSeparator + existingDyld; } catch { }
                logger.Debug($"Prepended LibreOffice program directory to PATH and DYLD_LIBRARY_PATH: {programDir}", verbose: true);
            }
            else
            {
                // Linux/Unix: set LD_LIBRARY_PATH as well
                var ldKey = "LD_LIBRARY_PATH";
                var existingLd = Environment.GetEnvironmentVariable(ldKey) ?? string.Empty;
                try { process.StartInfo.EnvironmentVariables[ldKey] = programDir + Path.PathSeparator + existingLd; } catch { }
                logger.Debug($"Prepended LibreOffice program directory to PATH and LD_LIBRARY_PATH: {programDir}", verbose: true);
            }
        }
        else
        {
            logger.Debug("Could not locate LibreOffice 'program' directory to adjust PATH/LD_LIBRARY_PATH.", verbose: true);
        }

        var stdout = new List<string>();
        var stderr = new List<string>();

        process.OutputDataReceived += (_, e) =>
        {
            if (e.Data != null) stdout.Add(e.Data);
        };
        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data != null) stderr.Add(e.Data);
        };

        try
        {
            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            // Wait with timeout (30 seconds should be plenty for IDL compilation)
            if (!process.WaitForExit(30000))
            {
                process.Kill();
                throw new TimeoutException("unoidl-write process timed out after 30 seconds.");
            }

            // Log any output
            foreach (var line in stdout)
            {
                logger.Debug(line, verbose: true);
            }

            if (process.ExitCode != 0)
            {
                var errorMessage = stderr.Count > 0
                    ? string.Join(Environment.NewLine, stderr)
                    : "No error details available.";

                throw new InvalidOperationException(
                    $"unoidl-write failed with exit code {process.ExitCode}.{Environment.NewLine}{errorMessage}");
            }

            // Log any warnings from stderr even on success
            foreach (var line in stderr)
            {
                logger.Warning(line);
            }

            logger.Success($"Generated {Path.GetFileName(rdbOutputPath)}");
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not TimeoutException)
        {
            throw new InvalidOperationException($"Failed to run unoidl-write: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Finds the types.rdb or offapi.rdb file in the SDK.
    /// </summary>
    private static string? FindTypesRdb(string sdkPath)
    {
        // Check various possible locations for the types registry
        var candidates = new[]
        {
            Path.Combine(sdkPath, "types.rdb"),
            Path.Combine(sdkPath, "idl", "types.rdb"),
            Path.Combine(sdkPath, "..", "program", "types.rdb"),
            Path.Combine(sdkPath, "..", "program", "types", "offapi.rdb"),
            // Linux paths
            Path.Combine(sdkPath, "..", "basis-link", "program", "offapi.rdb"),
            Path.Combine(sdkPath, "..", "ure-link", "share", "misc", "types.rdb"),
        };

        foreach (var candidate in candidates)
        {
            var fullPath = Path.GetFullPath(candidate);
            if (File.Exists(fullPath))
            {
                return fullPath;
            }
        }

        return null;
    }

    /// <summary>
    /// Attempts to locate the LibreOffice SDK installation path.
    /// </summary>
    public static string? GuessLibreOfficeSdkPath(Logger? logger = null)
    {
        // 1. Check environment variables first
        var envVars = new[] { "OO_SDK_HOME", "OO_SDK_PATH", "LIBREOFFICE_SDK_PATH", "LO_SDK_PATH" };
        foreach (var envVar in envVars)
        {
            var value = Environment.GetEnvironmentVariable(envVar);
            if (!string.IsNullOrEmpty(value) && Directory.Exists(value))
            {
                logger?.Debug($"Found SDK via {envVar}: {value}", verbose: true);
                return value;
            }
        }

        // 2. Platform-specific detection
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return GuessWindowsSdkPath(logger);
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            return GuessMacOsSdkPath(logger);
        }
        else
        {
            return GuessLinuxSdkPath(logger);
        }
    }

    private static string? GuessWindowsSdkPath(Logger? logger)
    {
        // Try registry first
        string? installPath = GetInstallPathFromRegistry();

        if (!string.IsNullOrEmpty(installPath))
        {
            // InstallPath usually points to \LibreOffice\program
            // SDK is at \LibreOffice\sdk
            var sdkPath = Path.GetFullPath(Path.Combine(installPath, "..", "sdk"));
            if (Directory.Exists(sdkPath))
            {
                logger?.Debug($"Found SDK via registry: {sdkPath}", verbose: true);
                return sdkPath;
            }
        }

        // Common installation paths
        var possiblePaths = new[]
        {
            @"C:\Program Files\LibreOffice\sdk",
            @"C:\Program Files (x86)\LibreOffice\sdk",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice", "sdk"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice", "sdk"),
        };

        foreach (var path in possiblePaths)
        {
            if (Directory.Exists(path))
            {
                logger?.Debug($"Found SDK at common path: {path}", verbose: true);
                return path;
            }
        }

        return null;
    }

    private static string? GetInstallPathFromRegistry()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return null;

        var subKeys = new[]
        {
            @"SOFTWARE\LibreOffice\UNO\InstallPath",
            @"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath"
        };

        var roots = new[] { Registry.CurrentUser, Registry.LocalMachine };

        foreach (var root in roots)
        {
            foreach (var subKey in subKeys)
            {
                try
                {
                    using var key = root.OpenSubKey(subKey);
                    var val = key?.GetValue("") as string;
                    if (!string.IsNullOrEmpty(val))
                        return val;
                }
                catch
                {
                    // Ignore registry access errors
                }
            }
        }

        return null;
    }

    private static string? GuessMacOsSdkPath(Logger? logger)
    {
        // 1. Try common Application paths first
        var possiblePaths = new[]
        {
            "/Applications/LibreOffice.app/Contents/sdk",
            "/Applications/LibreOffice.app/Contents/Resources/sdk",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Applications/LibreOffice.app/Contents/sdk"),
        };

        foreach (var path in possiblePaths)
        {
            if (Directory.Exists(path))
            {
                logger?.Debug($"Found SDK at: {path}", verbose: true);
                return path;
            }
        }

        // 2. Use 'which' to find libreoffice binary and derive SDK path
        var sdkFromWhich = FindSdkViaWhich(logger, "soffice", "libreoffice");
        if (sdkFromWhich != null)
            return sdkFromWhich;

        // 3. Try mdfind to locate LibreOffice.app on macOS
        var appPath = RunCommand("mdfind", "kMDItemKind == 'Application' && kMDItemFSName == 'LibreOffice.app'");
        if (!string.IsNullOrEmpty(appPath))
        {
            var firstPath = appPath.Split('\n').FirstOrDefault()?.Trim();
            if (!string.IsNullOrEmpty(firstPath))
            {
                var sdkPath = Path.Combine(firstPath, "Contents", "sdk");
                if (Directory.Exists(sdkPath))
                {
                    logger?.Debug($"Found SDK via mdfind: {sdkPath}", verbose: true);
                    return sdkPath;
                }
            }
        }

        return null;
    }

    private static string? GuessLinuxSdkPath(Logger? logger)
    {
        // 1. Use 'which' to find libreoffice binary and derive SDK path
        var sdkFromWhich = FindSdkViaWhich(logger, "libreoffice", "soffice", "lowriter");
        if (sdkFromWhich != null)
            return sdkFromWhich;

        // 2. Try to find LibreOffice installation via common paths
        var possiblePaths = new[]
        {
            "/usr/lib/libreoffice/sdk",
            "/usr/lib64/libreoffice/sdk",
            "/opt/libreoffice/sdk",
            "/opt/libreoffice7.0/sdk",
            "/opt/libreoffice7.1/sdk",
            "/opt/libreoffice7.2/sdk",
            "/opt/libreoffice7.3/sdk",
            "/opt/libreoffice7.4/sdk",
            "/opt/libreoffice7.5/sdk",
            "/opt/libreoffice7.6/sdk",
            "/opt/libreoffice24.2/sdk",
            "/opt/libreoffice24.8/sdk",
            "/usr/local/lib/libreoffice/sdk",
            "/snap/libreoffice/current/lib/libreoffice/sdk",
            // Flatpak paths
            "/var/lib/flatpak/app/org.libreoffice.LibreOffice/current/active/files/libreoffice/sdk",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".local/share/flatpak/app/org.libreoffice.LibreOffice/current/active/files/libreoffice/sdk"),
        };

        foreach (var path in possiblePaths)
        {
            if (Directory.Exists(path))
            {
                logger?.Debug($"Found SDK at: {path}", verbose: true);
                return path;
            }
        }

        // 3. Try to resolve from known binary paths directly
        var binaryPaths = new[] { "/usr/bin/libreoffice", "/usr/local/bin/libreoffice" };
        foreach (var binaryPath in binaryPaths)
        {
            var sdkPath = FindSdkFromBinary(binaryPath, logger);
            if (sdkPath != null)
                return sdkPath;
        }

        return null;
    }

    /// <summary>
    /// Uses 'which' command to find a binary and derive the SDK path from it.
    /// </summary>
    private static string? FindSdkViaWhich(Logger? logger, params string[] binaryNames)
    {
        foreach (var binaryName in binaryNames)
        {
            var whichOutput = RunCommand("which", binaryName);
            if (string.IsNullOrEmpty(whichOutput))
                continue;

            var binaryPath = whichOutput.Trim();
            if (!File.Exists(binaryPath))
                continue;

            logger?.Debug($"'which {binaryName}' returned: {binaryPath}", verbose: true);

            var sdkPath = FindSdkFromBinary(binaryPath, logger);
            if (sdkPath != null)
                return sdkPath;
        }

        return null;
    }

    /// <summary>
    /// Given a LibreOffice binary path, attempts to find the SDK directory.
    /// </summary>
    private static string? FindSdkFromBinary(string binaryPath, Logger? logger)
    {
        if (!File.Exists(binaryPath))
            return null;

        try
        {
            // Resolve symlinks to get the real path
            var realPath = ResolveSymlinkFully(binaryPath);
            if (realPath == null)
                return null;

            logger?.Debug($"Resolved binary path: {realPath}", verbose: true);

            // Try different relative paths from the binary to the SDK
            // Common structures:
            // - /usr/lib/libreoffice/program/soffice -> sdk is at /usr/lib/libreoffice/sdk
            // - /opt/libreoffice7.6/program/soffice -> sdk is at /opt/libreoffice7.6/sdk
            // - /Applications/LibreOffice.app/Contents/MacOS/soffice -> sdk is at Contents/sdk
            var dir = Path.GetDirectoryName(realPath);
            if (dir == null)
                return null;

            var relativePaths = new[]
            {
                Path.Combine(dir, "..", "sdk"),           // program/soffice -> sdk
                Path.Combine(dir, "..", "..", "sdk"),     // MacOS/soffice -> Contents/sdk
                Path.Combine(dir, "sdk"),                 // Direct sibling
            };

            foreach (var relPath in relativePaths)
            {
                var sdkPath = Path.GetFullPath(relPath);
                if (Directory.Exists(sdkPath))
                {
                    logger?.Debug($"Found SDK via binary: {sdkPath}", verbose: true);
                    return sdkPath;
                }
            }
        }
        catch
        {
            // Ignore errors in path resolution
        }

        return null;
    }

    /// <summary>
    /// Resolves a symlink to its immediate target path (one level).
    /// </summary>
    private static string? ResolveSymlink(string path)
    {
        try
        {
            var fileInfo = new FileInfo(path);
            if (fileInfo.LinkTarget != null)
            {
                // If it's a relative link, resolve it
                if (!Path.IsPathRooted(fileInfo.LinkTarget))
                {
                    var dir = Path.GetDirectoryName(path);
                    return dir != null
                        ? Path.GetFullPath(Path.Combine(dir, fileInfo.LinkTarget))
                        : null;
                }
                return fileInfo.LinkTarget;
            }
            return path;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Resolves a symlink fully, following all links to the final target.
    /// </summary>
    private static string? ResolveSymlinkFully(string path)
    {
        try
        {
            // Use readlink -f on Linux/macOS for full resolution
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                var realPath = RunCommand("readlink", $"-f \"{path}\"");
                if (!string.IsNullOrEmpty(realPath))
                {
                    var resolved = realPath.Trim();
                    if (File.Exists(resolved))
                        return resolved;
                }

                // Fallback: try 'realpath' command
                realPath = RunCommand("realpath", $"\"{path}\"");
                if (!string.IsNullOrEmpty(realPath))
                {
                    var resolved = realPath.Trim();
                    if (File.Exists(resolved))
                        return resolved;
                }
            }

            // Manual resolution as fallback
            var current = path;
            var visited = new HashSet<string>(StringComparer.Ordinal);
            const int maxDepth = 20; // Prevent infinite loops

            for (int i = 0; i < maxDepth; i++)
            {
                if (!visited.Add(current))
                    break; // Circular symlink detected

                var resolved = ResolveSymlink(current);
                if (resolved == null || resolved == current)
                    break;

                current = resolved;
            }

            return File.Exists(current) ? current : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Runs a command and returns its stdout output, or null on failure.
    /// </summary>
    private static string? RunCommand(string command, string arguments)
    {
        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = command,
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                }
            };

            process.Start();

            var output = process.StandardOutput.ReadToEnd();

            if (!process.WaitForExit(5000))
            {
                try { process.Kill(); } catch { }
                return null;
            }

            return process.ExitCode == 0 ? output : null;
        }
        catch
        {
            return null;
        }
    }
}
