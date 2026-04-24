using System.IO;

namespace Envoy.Services;

/// <summary>
/// Central source of truth for the per-user AppData folder.
/// Also one-time-migrates data from the old "CavtEmail" folder (the app's
/// previous name) so existing users don't lose their contacts or config.
/// </summary>
internal static class AppDataPaths
{
    private const string CurrentFolder = "Envoy";
    private const string LegacyFolder  = "CavtEmail";

    private static readonly object _lock = new();
    private static bool _migrated;

    public static string Dir
    {
        get
        {
            var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var current = Path.Combine(root, CurrentFolder);
            EnsureMigrated(root, current);
            Directory.CreateDirectory(current);
            return current;
        }
    }

    private static void EnsureMigrated(string root, string current)
    {
        if (_migrated) return;
        lock (_lock)
        {
            if (_migrated) return;
            _migrated = true;

            var legacy = Path.Combine(root, LegacyFolder);
            if (!Directory.Exists(legacy) || Directory.Exists(current)) return;

            try
            {
                // Simple case: rename the whole folder.
                Directory.Move(legacy, current);
            }
            catch
            {
                // Fallback: copy each file across; leave the legacy folder intact.
                try
                {
                    Directory.CreateDirectory(current);
                    foreach (var src in Directory.EnumerateFiles(legacy))
                    {
                        var dst = Path.Combine(current, Path.GetFileName(src));
                        if (!File.Exists(dst)) File.Copy(src, dst);
                    }
                }
                catch { /* best effort */ }
            }
        }
    }
}
