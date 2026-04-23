using System.Runtime.InteropServices;
using System.Text;

namespace CavtEmail.Services;

/// <summary>
/// Looks up the current Windows user's display name via Secur32 — works on
/// domain-joined machines (returns AD display name) and returns null
/// (so the caller can fall back) otherwise.
/// </summary>
public static class CurrentUserInfo
{
    // EXTENDED_NAME_FORMAT values
    private const int NameDisplay = 3;
    private const int NameUserPrincipal = 8;

    [DllImport("secur32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool GetUserNameEx(int nameFormat, StringBuilder userName, ref uint size);

    public static string? TryGetDisplayName()
    {
        var s = Query(NameDisplay);
        return string.IsNullOrWhiteSpace(s) ? null : s;
    }

    /// <summary>Typically 'first.last@domain' on domain-joined machines. May be null.</summary>
    public static string? TryGetUpn()
    {
        var s = Query(NameUserPrincipal);
        return string.IsNullOrWhiteSpace(s) ? null : s;
    }

    private static string? Query(int format)
    {
        try
        {
            uint size = 256;
            var sb = new StringBuilder((int)size);
            if (GetUserNameEx(format, sb, ref size))
                return sb.ToString();

            // ERROR_MORE_DATA = 234 → retry with reported size
            if (Marshal.GetLastWin32Error() == 234)
            {
                sb = new StringBuilder((int)size);
                if (GetUserNameEx(format, sb, ref size))
                    return sb.ToString();
            }
        }
        catch
        {
            // fall through
        }
        return null;
    }
}
