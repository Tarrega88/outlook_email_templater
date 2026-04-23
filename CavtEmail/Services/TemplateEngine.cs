using System.Text.RegularExpressions;
using CavtEmail.Models;

namespace CavtEmail.Services;

/// <summary>
/// Replaces {{var}} tokens in subject/body strings.
///
/// Available tokens:
///   {{recipients}}        comma-separated To display names (falls back to email)
///   {{recipientEmails}}   comma-separated To emails
///   {{recipientCount}}    number of To recipients
///   {{greetRecipients}}   To recipients' first names with oxford-comma ("Alice, Bob, and Carol")
///   {{cc}}                comma-separated Cc display names
///   {{ccEmails}}          comma-separated Cc emails
///   {{files}}             newline-separated file names
///   {{fileList}}          bulleted file list (" - name")
///   {{fileCount}}
///   {{sender}}            global sender name
///   {{date}}              today, long format
///   {{time}}              current time
///   {{group}}             group name
/// </summary>
public static class TemplateEngine
{
    private static readonly Regex Token = new(@"\{\{\s*([A-Za-z0-9_.]+)\s*\}\}", RegexOptions.Compiled);

    public static string Render(string template, EmailGroup group, AppConfig config)
    {
        if (string.IsNullOrEmpty(template)) return template ?? "";

        return Token.Replace(template, m =>
        {
            var key = m.Groups[1].Value;
            return Resolve(key, group, config) ?? m.Value;
        });
    }

    private static string? Resolve(string key, EmailGroup group, AppConfig config)
    {
        var toList = group.Recipients.Where(r => !r.IsCc).ToList();
        var ccList = group.Recipients.Where(r =>  r.IsCc).ToList();

        switch (key.ToLowerInvariant())
        {
            case "recipients":
                return string.Join(", ", toList.Select(r =>
                    string.IsNullOrWhiteSpace(r.Name) ? r.Email : r.Name));
            case "recipientemails":
                return string.Join(", ", toList.Select(r => r.Email));
            case "recipientcount":
                return toList.Count.ToString();
            case "greetrecipients":
                return NameUtil.JoinOxford(toList.Select(r =>
                    NameUtil.FirstName(r.Name) ?? r.Email));
            case "cc":
                return string.Join(", ", ccList.Select(r =>
                    string.IsNullOrWhiteSpace(r.Name) ? r.Email : r.Name));
            case "ccemails":
                return string.Join(", ", ccList.Select(r => r.Email));
            case "files":
                return string.Join(Environment.NewLine, ResolvedFileNames(group));
            case "filelist":
                return string.Join(Environment.NewLine, ResolvedFileNames(group).Select(n => " - " + n));
            case "filecount":
                return ResolvedFileNames(group).Count().ToString();
            case "sender":
                return config.SenderName;
            case "date":
                return DateTime.Now.ToLongDateString();
            case "time":
                return DateTime.Now.ToShortTimeString();
            case "group":
                return group.Name;
        }

        return null;
    }

    /// <summary>
    /// The file names actually being attached for a group, after expanding folder
    /// references to their top-level files.
    /// </summary>
    private static IEnumerable<string> ResolvedFileNames(EmailGroup group) =>
        group.Attachments
            .SelectMany(a => a.ResolveFiles())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(System.IO.Path.GetFileName)!;
}
