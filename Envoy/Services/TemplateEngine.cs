using System.Text.RegularExpressions;
using Envoy.Models;

namespace Envoy.Services;

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
///   {{email}}             this email's template name (alias: {{group}} for older configs)
/// </summary>
public static class TemplateEngine
{
    private static readonly Regex Token = new(@"\{\{\s*([A-Za-z0-9_.]+)\s*\}\}", RegexOptions.Compiled);

    public static string Render(string template, EmailTemplate email, AppConfig config)
    {
        if (string.IsNullOrEmpty(template)) return template ?? "";

        return Token.Replace(template, m =>
        {
            var key = m.Groups[1].Value;
            return Resolve(key, email, config) ?? m.Value;
        });
    }

    private static string? Resolve(string key, EmailTemplate email, AppConfig config)
    {
        var toList = email.Recipients.Where(r => !r.IsCc).ToList();
        var ccList = email.Recipients.Where(r =>  r.IsCc).ToList();

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
                return string.Join(Environment.NewLine, ResolvedFileNames(email));
            case "filelist":
                return string.Join(Environment.NewLine, ResolvedFileNames(email).Select(n => " - " + n));
            case "filecount":
                return ResolvedFileNames(email).Count().ToString();
            case "sender":
                return config.SenderName;
            case "date":
                return DateTime.Now.ToLongDateString();
            case "time":
                return DateTime.Now.ToShortTimeString();
            case "email":
            case "group": // legacy alias
                return email.Name;
        }

        return null;
    }

    /// <summary>
    /// The file names actually being attached for an email, after expanding folder
    /// references to their top-level files.
    /// </summary>
    private static IEnumerable<string> ResolvedFileNames(EmailTemplate email) =>
        email.Attachments
            .SelectMany(a => a.ResolveFiles())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(System.IO.Path.GetFileName)!;
}
