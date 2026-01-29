using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;

namespace CalcDNA.CLI;

/// <summary>
/// Generates RDB (Registry Database) files from IDL using LibreOffice SDK tools.
/// RDB files contain compiled type information used by the UNO runtime.
/// </summary>
internal static class RdbGenerator
{
    private const int ProcessTimeoutMs = 60000; // 60 seconds

    /// <summary>
    /// Generates an RDB file from an IDL file using the LibreOffice SDK's unoidl-write tool.
    /// </summary>
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

        var binDir = Path.Combine(sdkActualPath, "bin");
        var unoidlWrite = Path.Combine(binDir, exeName);

        if (!File.Exists(unoidlWrite))
        {
            // Try sibling bin directory if not found in sdk/bin
            unoidlWrite = Path.Combine(sdkActualPath, "..", "bin", exeName);
            if (!File.Exists(unoidlWrite))
            {
                throw new FileNotFoundException(
                    $"unoidl-write not found in the LibreOffice SDK bin directory. Expected at: {unoidlWrite}",
                    unoidlWrite);
            }
        }

        // Locate required type definition files
        var typesRdb = FindTypesRdb(sdkActualPath, logger);

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

        ConfigureEnvironment(process.StartInfo, sdkActualPath, logger);

        var stdout = new List<string>();
        var stderr = new List<string>();

        process.OutputDataReceived += (_, e) => { if (e.Data != null) stdout.Add(e.Data); };
        process.ErrorDataReceived += (_, e) => { if (e.Data != null) stderr.Add(e.Data); };

        try
        {
            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            if (!process.WaitForExit(ProcessTimeoutMs))
            {
                try { process.Kill(true); } catch { }
                throw new TimeoutException($"unoidl-write process timed out after {ProcessTimeoutMs/1000} seconds.");
            }

            foreach (var line in stdout) logger.Debug(line, verbose: true);

            if (process.ExitCode != 0)
            {
                var errorMessage = stderr.Count > 0
                    ? string.Join(Environment.NewLine, stderr)
                    : "No error details available from stderr.";

                throw new InvalidOperationException(
                    $"unoidl-write failed with exit code {process.ExitCode}.{Environment.NewLine}{errorMessage}");
            }

            foreach (var line in stderr) logger.Warning(line);

            logger.Success($"Generated {Path.GetFileName(rdbOutputPath)}");
        }
        catch (Exception ex) when (ex is not InvalidOperationException and not TimeoutException)
        {
            throw new InvalidOperationException($"Failed to run unoidl-write: {ex.Message}", ex);
        }
    }

    private static void ConfigureEnvironment(ProcessStartInfo startInfo, string sdkPath, Logger logger)
    {
        var programDirCandidates = new[]
        {
            Path.Combine(sdkPath, "..", "program"),
            Path.Combine(sdkPath, "program"),
            Path.Combine(sdkPath, "..", "bin"),
            sdkPath
        };

        string? programDir = null;
        foreach (var cand in programDirCandidates)
        {
            try
            {
                if (Directory.Exists(cand))
                {
                    programDir = Path.GetFullPath(cand);
                    break;
                }
            } catch { }
        }

        if (programDir != null)
        {
            string pathVar = "PATH";
            string currentPath = Environment.GetEnvironmentVariable(pathVar) ?? "";
            startInfo.EnvironmentVariables[pathVar] = programDir + Path.PathSeparator + currentPath;

            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                string dyldVar = "DYLD_LIBRARY_PATH";
                string currentDyld = Environment.GetEnvironmentVariable(dyldVar) ?? "";
                startInfo.EnvironmentVariables[dyldVar] = programDir + Path.PathSeparator + currentDyld;
            }
            else if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                string ldVar = "LD_LIBRARY_PATH";
                string currentLd = Environment.GetEnvironmentVariable(ldVar) ?? "";
                startInfo.EnvironmentVariables[ldVar] = programDir + Path.PathSeparator + currentLd;
            }
        }
    }

    private static string? FindTypesRdb(string sdkPath, Logger logger)
    {
        var candidates = new[]
        {
            Path.Combine(sdkPath, "types.rdb"),
            Path.Combine(sdkPath, "idl", "types.rdb"),
            Path.Combine(sdkPath, "..", "program", "types.rdb"),
            Path.Combine(sdkPath, "..", "program", "types", "offapi.rdb"),
            Path.Combine(sdkPath, "..", "basis-link", "program", "offapi.rdb"),
            Path.Combine(sdkPath, "..", "ure-link", "share", "misc", "types.rdb"),
        };

        foreach (var candidate in candidates)
        {
            try
            {
                var fullPath = Path.GetFullPath(candidate);
                if (File.Exists(fullPath)) return fullPath;
            } catch { }
        }

        return null;
    }

    public static string? GuessLibreOfficeSdkPath(Logger? logger = null)
    {
        var envVars = new[] { "OO_SDK_HOME", "OO_SDK_PATH", "LIBREOFFICE_SDK_PATH", "LO_SDK_PATH" };
        foreach (var envVar in envVars)
        {
            var value = Environment.GetEnvironmentVariable(envVar);
            if (!string.IsNullOrEmpty(value) && Directory.Exists(value)) return value;
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return GuessWindowsSdkPath(logger);
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return GuessMacOsSdkPath(logger);
        return GuessLinuxSdkPath(logger);
    }

    private static string? GuessWindowsSdkPath(Logger? logger)
    {
        string? installPath = GetInstallPathFromRegistry();
        if (!string.IsNullOrEmpty(installPath))
        {
            var sdkPath = Path.GetFullPath(Path.Combine(installPath, "..", "sdk"));
            if (Directory.Exists(sdkPath)) return sdkPath;
        }

        var common = new[]
        {
            @"C:\Program Files\LibreOffice\sdk",
            @"C:\Program Files (x86)\LibreOffice\sdk",
        };

        return common.FirstOrDefault(Directory.Exists);
    }

    private static string? GetInstallPathFromRegistry()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return null;

        var paths = new[] { @"SOFTWARE\LibreOffice\UNO\InstallPath", @"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath" };
        var roots = new[] { Registry.CurrentUser, Registry.LocalMachine };

        foreach (var root in roots)
        {
            foreach (var path in paths)
            {
                try
                {
                    using var key = root.OpenSubKey(path);
                    var val = key?.GetValue("") as string;
                    if (!string.IsNullOrEmpty(val)) return val;
                } catch { }
            }
        }
        return null;
    }

    private static string? GuessMacOsSdkPath(Logger? logger)
    {
        var paths = new[]
        {
            "/Applications/LibreOffice.app/Contents/sdk",
            "/Applications/LibreOffice.app/Contents/Resources/sdk",
        };

        var existing = paths.FirstOrDefault(Directory.Exists);
        if (existing != null) return existing;

        var appPath = RunCommand("mdfind", "kMDItemKind == 'Application' && kMDItemFSName == 'LibreOffice.app'");
        if (!string.IsNullOrEmpty(appPath))
        {
            var first = appPath.Split('\n').FirstOrDefault()?.Trim();
            if (!string.IsNullOrEmpty(first))
            {
                var sdk = Path.Combine(first, "Contents", "sdk");
                if (Directory.Exists(sdk)) return sdk;
            }
        }

        return FindSdkViaWhich(logger, "soffice");
    }

    private static string? GuessLinuxSdkPath(Logger? logger)
    {
        var paths = new[]
        {
            "/usr/lib/libreoffice/sdk",
            "/usr/lib64/libreoffice/sdk",
            "/opt/libreoffice/sdk",
            "/usr/local/lib/libreoffice/sdk",
        };

        var existing = paths.FirstOrDefault(Directory.Exists);
        if (existing != null) return existing;

        return FindSdkViaWhich(logger, "libreoffice", "soffice");
    }

    private static string? FindSdkViaWhich(Logger? logger, params string[] binaryNames)
    {
        foreach (var name in binaryNames)
        {
            var output = RunCommand("which", name);
            if (string.IsNullOrEmpty(output)) continue;

            var path = output.Trim();
            var realPath = ResolveSymlinkFully(path);
            if (realPath == null) continue;

            var dir = Path.GetDirectoryName(realPath);
            if (dir == null) continue;

            var relatives = new[] { Path.Combine(dir, "..", "sdk"), Path.Combine(dir, "..", "..", "sdk"), Path.Combine(dir, "sdk") };
            foreach (var rel in relatives)
            {
                try {
                    var full = Path.GetFullPath(rel);
                    if (Directory.Exists(full)) return full;
                } catch { }
            }
        }
        return null;
    }

    private static string? ResolveSymlinkFully(string path)
    {
        if (!File.Exists(path)) return null;

        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            var output = RunCommand("readlink", $"-f \"{path}\"") ?? RunCommand("realpath", $"\"{path}\"");
            if (!string.IsNullOrEmpty(output))
            {
                var resolved = output.Trim();
                if (File.Exists(resolved)) return resolved;
            }
        }

        var current = path;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < 20; i++)
        {
            if (!visited.Add(current)) break;
            try {
                var info = new FileInfo(current);
                if (info.LinkTarget == null) return current;
                current = Path.IsPathRooted(info.LinkTarget) ? info.LinkTarget : Path.GetFullPath(Path.Combine(Path.GetDirectoryName(current)!, info.LinkTarget));
            } catch { break; }
        }
        return current;
    }

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
            if (process.WaitForExit(5000) && process.ExitCode == 0) return output;
        } catch { }
        return null;
    }
}
