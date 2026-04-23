namespace CavtEmail.Services;

using System.Linq;

/// <summary>
/// Splits a human name into first / last parts using whitespace.
///
/// Rules (per app spec):
///   FirstName  -> first whitespace-separated token, only if the name has any text.
///   LastName   -> last whitespace-separated token, only if there are 2+ tokens
///                 (so "Michael See" yields "See" but "Michael" yields null).
/// Middle names/initials are treated as part of neither; the last token is the last name.
/// </summary>
public static class NameUtil
{
    public static string? FirstName(string? fullName)
    {
        if (string.IsNullOrWhiteSpace(fullName)) return null;
        var parts = fullName.Trim().Split(new[] { ' ', '\t' }, System.StringSplitOptions.RemoveEmptyEntries);
        return parts.Length == 0 ? null : parts[0];
    }

    public static string? LastName(string? fullName)
    {
        if (string.IsNullOrWhiteSpace(fullName)) return null;
        var parts = fullName.Trim().Split(new[] { ' ', '\t' }, System.StringSplitOptions.RemoveEmptyEntries);
        return parts.Length < 2 ? null : parts[^1];
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
