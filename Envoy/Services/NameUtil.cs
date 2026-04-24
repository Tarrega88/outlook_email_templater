namespace Envoy.Services;

using System.Linq;
using System.Text.RegularExpressions;

/// <summary>
/// Splits a human name into first / last parts.
///
/// Handles the two common shapes:
///   "First Middle Last"            -> First, Last
///   "Last, First M."               -> First, Last  (Outlook GAL format)
/// Trailing parentheticals like "(It Concepts, Inc.)" and trailing single-letter
/// initials with optional periods are stripped.
/// </summary>
public static class NameUtil
{
    private static readonly Regex Parenthetical = new(@"\s*\([^)]*\)\s*$", RegexOptions.Compiled);

    /// <summary>
    /// Normalize raw display name to "First Last" order, parentheticals removed.
    /// Returns null when there's nothing usable.
    /// </summary>
    private static string? Normalize(string? fullName)
    {
        if (string.IsNullOrWhiteSpace(fullName)) return null;
        var s = Parenthetical.Replace(fullName.Trim(), "").Trim();
        if (s.Length == 0) return null;

        // "Last, First M." -> "First M. Last"
        var commaIdx = s.IndexOf(',');
        if (commaIdx > 0)
        {
            var last = s[..commaIdx].Trim();
            var rest = s[(commaIdx + 1)..].Trim();
            if (last.Length > 0 && rest.Length > 0) s = $"{rest} {last}";
        }
        return s;
    }

    public static string? FirstName(string? fullName)
    {
        var s = Normalize(fullName);
        if (s == null) return null;
        var parts = s.Split(new[] { ' ', '\t' }, System.StringSplitOptions.RemoveEmptyEntries);
        return parts.Length == 0 ? null : parts[0].TrimEnd('.', ',');
    }

    public static string? LastName(string? fullName)
    {
        var s = Normalize(fullName);
        if (s == null) return null;
        var parts = s.Split(new[] { ' ', '\t' }, System.StringSplitOptions.RemoveEmptyEntries)
                     .Select(p => p.TrimEnd('.', ','))
                     .Where(p => p.Length > 0)
                     .ToArray();
        if (parts.Length < 2) return null;

        // Drop trailing single-letter initials (e.g., "Michael W" -> "Michael").
        // We want the actual surname — the last non-initial token.
        for (int i = parts.Length - 1; i >= 1; i--)
        {
            if (parts[i].Length > 1) return parts[i];
        }
        return parts[^1];
    }

    /// <summary>
    /// Joins a sequence of names with an oxford comma:
    ///   ["Michael"]                    -> "Michael"
    ///   ["Michael","Bob"]              -> "Michael and Bob"
    ///   ["Michael","Bob","Tim"]        -> "Michael, Bob, and Tim"
    /// Null/whitespace entries are skipped.
    /// </summary>
    public static string JoinOxford(System.Collections.Generic.IEnumerable<string?> names)
    {
        var list = new System.Collections.Generic.List<string>();
        foreach (var n in names)
            if (!string.IsNullOrWhiteSpace(n)) list.Add(n!.Trim());

        return list.Count switch
        {
            0 => "",
            1 => list[0],
            2 => $"{list[0]} and {list[1]}",
            _ => string.Join(", ", list.Take(list.Count - 1)) + ", and " + list[^1],
        };
    }
}
