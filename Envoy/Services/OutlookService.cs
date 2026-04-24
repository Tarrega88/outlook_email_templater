using System.IO;
using System.Runtime.InteropServices;

namespace Envoy.Services;

/// <summary>
/// Late-bound Outlook automation. Avoids any compile-time reference to the Outlook
/// interop assembly — the user just needs classic Outlook installed at runtime.
/// </summary>
public static class OutlookService
{
    // Outlook MailItem.BodyFormat values
    private const int OlFormatPlain = 1;
    private const int OlFormatHtml = 2;

    // Outlook OlItemType.olMailItem
    private const int OlMailItem = 0;

    public enum Mode { Draft, Send }

    public static void CreateMail(
        IEnumerable<string> toEmails,
        IEnumerable<string> ccEmails,
        string subject,
        string body,
        IEnumerable<string> attachmentPaths,
        bool htmlBody,
        Mode mode)
    {
        var appType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException(
                "Could not find classic Outlook on this machine (Outlook.Application ProgID missing).");

        dynamic? outlook = null;
        dynamic? mail = null;
        try
        {
            // Outlook.Application is a singleton — Activator returns the running
            // instance if Outlook is already open, otherwise it starts one.
            outlook = Activator.CreateInstance(appType)
                ?? throw new InvalidOperationException("Failed to start Outlook.");

            mail = outlook.CreateItem(OlMailItem);

            mail.Subject = subject ?? "";
            mail.BodyFormat = htmlBody ? OlFormatHtml : OlFormatPlain;
            if (htmlBody) mail.HTMLBody = body ?? "";
            else mail.Body = body ?? "";

            var to = string.Join("; ", toEmails.Where(s => !string.IsNullOrWhiteSpace(s)));
            if (!string.IsNullOrWhiteSpace(to))
                mail.To = to;

            var cc = string.Join("; ", ccEmails.Where(s => !string.IsNullOrWhiteSpace(s)));
            if (!string.IsNullOrWhiteSpace(cc))
                mail.CC = cc;

            foreach (var path in attachmentPaths)
            {
                if (string.IsNullOrWhiteSpace(path)) continue;
                if (!File.Exists(path))
                    throw new FileNotFoundException("Attachment not found", path);
                mail.Attachments.Add(path);
            }

            if (mode == Mode.Send)
                mail.Send();
            else
                mail.Display(false); // non-modal
        }
        finally
        {
            if (mail != null) Marshal.FinalReleaseComObject(mail);
            if (outlook != null) Marshal.FinalReleaseComObject(outlook);
        }
    }
}
