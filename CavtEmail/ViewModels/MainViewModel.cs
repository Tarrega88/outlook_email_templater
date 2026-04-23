using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using CavtEmail.Models;
using CavtEmail.Services;
using Microsoft.Win32;

namespace CavtEmail.ViewModels;

public class MainViewModel : NotifyBase
{
    private AppConfig _config = new();
    private EmailGroup? _selectedGroup;
    private string _currentConfigPath;
    private string _statusMessage = "";

    public MainViewModel()
    {
        _currentConfigPath = ConfigService.DefaultPath;
        _config = ConfigService.Load(_currentConfigPath);

        // Contacts live in a shared file independent of whatever config is open.
        _config.Contacts = ContactsService.Load();
        MigrateLegacyContacts(_config);

        if (_config.Groups.Count == 0)
        {
            _config.Groups.Add(new EmailGroup { Name = "Group 1" });
        }
        _selectedGroup = _config.Groups.FirstOrDefault();

        // First-run: auto-fill the sender display name from Windows / AD.
        if (string.IsNullOrWhiteSpace(_config.SenderName))
        {
            _config.SenderName = CurrentUserInfo.TryGetDisplayName()
                                 ?? Environment.UserName
                                 ?? "";
        }

        AddGroupCommand          = new RelayCommand(AddGroup);
        RemoveGroupCommand       = new RelayCommand(RemoveGroup, () => SelectedGroup != null);
        DuplicateGroupCommand    = new RelayCommand(DuplicateGroup, () => SelectedGroup != null);

        AddRecipientCommand      = new RelayCommand(AddRecipient, () => SelectedGroup != null);
        RemoveRecipientCommand   = new RelayCommand(RemoveRecipient);

        AddAttachmentCommand        = new RelayCommand(AddAttachment, () => SelectedGroup != null);
        AddAttachmentFolderCommand  = new RelayCommand(AddAttachmentFolder, () => SelectedGroup != null);
        RemoveAttachmentCommand     = new RelayCommand(RemoveAttachment);

        SaveCommand              = new RelayCommand(Save);
        SaveAsCommand            = new RelayCommand(SaveAs);
        OpenCommand              = new RelayCommand(Open);

        SendCommand              = new RelayCommand(Send, () => SelectedGroup != null);
        SendAllCommand           = new RelayCommand(SendAll, () => _config.Groups.Count > 0);
        PreviewCommand           = new RelayCommand(Preview, () => SelectedGroup != null);
        ManageContactsCommand    = new RelayCommand(ManageContacts);
    }

    public AppConfig Config { get => _config; set { Set(ref _config, value); Raise(nameof(Groups)); } }
    public ObservableCollection<EmailGroup> Groups => _config.Groups;

    public EmailGroup? SelectedGroup
    {
        get => _selectedGroup;
        set { if (Set(ref _selectedGroup, value)) Raise(nameof(HasSelection)); }
    }

    public bool HasSelection => _selectedGroup != null;

    public string StatusMessage { get => _statusMessage; set => Set(ref _statusMessage, value); }

    public string CurrentConfigPath { get => _currentConfigPath; set => Set(ref _currentConfigPath, value); }

    public Array SendModes { get; } = Enum.GetValues(typeof(SendMode));

    public RelayCommand AddGroupCommand { get; }
    public RelayCommand RemoveGroupCommand { get; }
    public RelayCommand DuplicateGroupCommand { get; }
    public RelayCommand AddRecipientCommand { get; }
    public RelayCommand RemoveRecipientCommand { get; }
    public RelayCommand AddAttachmentCommand { get; }
    public RelayCommand AddAttachmentFolderCommand { get; }
    public RelayCommand RemoveAttachmentCommand { get; }
    public RelayCommand SaveCommand { get; }
    public RelayCommand SaveAsCommand { get; }
    public RelayCommand OpenCommand { get; }
    public RelayCommand SendCommand { get; }
    public RelayCommand SendAllCommand { get; }
    public RelayCommand PreviewCommand { get; }
    public RelayCommand ManageContactsCommand { get; }

    public ObservableCollection<Contact> Contacts => _config.Contacts;

    /// <summary>Add any recipient (name,email) pair that isn't already in the contact book.</summary>
    public void CaptureContactsFrom(EmailGroup g)
    {
        var added = ContactsService.Merge(_config.Contacts,
            g.Recipients.Select(r => new Contact { Name = r.Name, Email = r.Email }));
        if (added > 0) SaveContacts();
    }

    /// <summary>Persist the shared contacts file. Safe to call freely.</summary>
    public void SaveContacts()
    {
        try { ContactsService.Save(_config.Contacts); }
        catch (Exception ex) { StatusMessage = "Contacts save failed: " + ex.Message; }
    }

    /// <summary>
    /// If an opened config carried an embedded "Contacts" array (old format),
    /// fold those entries into the shared store and persist.
    /// </summary>
    private void MigrateLegacyContacts(AppConfig cfg)
    {
        if (cfg.LegacyContacts == null || cfg.LegacyContacts.Count == 0) return;
        var added = ContactsService.Merge(cfg.Contacts, cfg.LegacyContacts);
        cfg.LegacyContacts = null;
        if (added > 0) SaveContacts();
    }

    private void ManageContacts()
    {
        var win = new ContactsWindow(_config.Contacts) { Owner = Application.Current.MainWindow };
        win.ShowDialog();
        SaveContacts();
    }

    private void AddGroup()
    {
        var g = new EmailGroup { Name = $"Group {_config.Groups.Count + 1}" };
        _config.Groups.Add(g);
        SelectedGroup = g;
    }

    private void RemoveGroup()
    {
        if (SelectedGroup == null) return;
        if (MessageBox.Show($"Delete group '{SelectedGroup.Name}'?", "Confirm",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
        var idx = _config.Groups.IndexOf(SelectedGroup);
        _config.Groups.Remove(SelectedGroup);
        SelectedGroup = _config.Groups.Count == 0
            ? null
            : _config.Groups[Math.Min(idx, _config.Groups.Count - 1)];
    }

    private void DuplicateGroup()
    {
        if (SelectedGroup == null) return;
        var src = SelectedGroup;
        var copy = new EmailGroup
        {
            Name = src.Name + " (copy)",
            Subject = src.Subject,
            Body = src.Body,
        };
        foreach (var r in src.Recipients) copy.Recipients.Add(new Recipient { Name = r.Name, Email = r.Email, IsCc = r.IsCc });
        foreach (var a in src.Attachments) copy.Attachments.Add(new AttachmentItem { Path = a.Path });
        _config.Groups.Add(copy);
        SelectedGroup = copy;
    }

    private void AddRecipient()
    {
        SelectedGroup?.Recipients.Add(new Recipient { Name = "", Email = "" });
    }

    private void RemoveRecipient(object? param)
    {
        if (SelectedGroup != null && param is Recipient r)
            SelectedGroup.Recipients.Remove(r);
    }

    private void AddAttachment()
    {
        if (SelectedGroup == null) return;
        var dlg = new OpenFileDialog { Multiselect = true, Title = "Select files to attach" };
        if (dlg.ShowDialog() == true)
        {
            foreach (var f in dlg.FileNames)
                AddAttachmentPath(f);
        }
    }

    private void AddAttachmentFolder()
    {
        if (SelectedGroup == null) return;
        var dlg = new OpenFolderDialog
        {
            Title = "Select a folder — its top-level files will be attached at send time",
            Multiselect = true
        };
        if (dlg.ShowDialog() == true)
        {
            foreach (var f in dlg.FolderNames)
                AddAttachmentPath(f);
        }
    }

    public void AddAttachmentPath(string path)
    {
        if (SelectedGroup == null || string.IsNullOrWhiteSpace(path)) return;
        if (SelectedGroup.Attachments.Any(a => string.Equals(a.Path, path, StringComparison.OrdinalIgnoreCase))) return;
        SelectedGroup.Attachments.Add(new AttachmentItem { Path = path });
    }

    private void RemoveAttachment(object? param)
    {
        if (SelectedGroup != null && param is AttachmentItem a)
            SelectedGroup.Attachments.Remove(a);
    }

    private void Save()
    {
        try
        {
            ConfigService.Save(_config, _currentConfigPath);
            StatusMessage = $"Saved to {_currentConfigPath}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Save failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveAs()
    {
        var dlg = new SaveFileDialog
        {
            Filter = "JSON config (*.json)|*.json|All files (*.*)|*.*",
            FileName = Path.GetFileName(_currentConfigPath),
            InitialDirectory = Path.GetDirectoryName(_currentConfigPath)
        };
        if (dlg.ShowDialog() == true)
        {
            CurrentConfigPath = dlg.FileName;
            Save();
        }
    }

    private void Open()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "JSON config (*.json)|*.json|All files (*.*)|*.*",
            InitialDirectory = Path.GetDirectoryName(_currentConfigPath)
        };
        if (dlg.ShowDialog() == true)
        {
            try
            {
                var cfg = ConfigService.Load(dlg.FileName);
                // Preserve the shared contacts list across the swap.
                cfg.Contacts = _config.Contacts;
                MigrateLegacyContacts(cfg);
                Config = cfg;
                CurrentConfigPath = dlg.FileName;
                SelectedGroup = cfg.Groups.FirstOrDefault();
                StatusMessage = $"Loaded {dlg.FileName}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Open failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void Preview()
    {
        if (SelectedGroup == null) return;
        var subject = TemplateEngine.Render(SelectedGroup.Subject, SelectedGroup, _config);
        var body = TemplateEngine.Render(SelectedGroup.Body, SelectedGroup, _config);
        var to = string.Join("; ", SelectedGroup.Recipients.Where(r => !r.IsCc).Select(r => r.Email));
        var cc = string.Join("; ", SelectedGroup.Recipients.Where(r =>  r.IsCc).Select(r => r.Email));
        var files = string.Join(Environment.NewLine,
            SelectedGroup.Attachments.SelectMany(a =>
                a.IsFolder
                    ? new[] { $" - [folder] {a.Path}" }
                        .Concat(a.ResolveFiles().Select(f => "     • " + f))
                    : new[] { " - " + a.Path }));
        var ccLine = string.IsNullOrWhiteSpace(cc) ? "" : $"Cc: {cc}\r\n";
        var msg = $"To: {to}\r\n{ccLine}Subject: {subject}\r\n\r\nAttachments:\r\n{files}\r\n\r\n---\r\n{body}";
        MessageBox.Show(msg, "Preview: " + SelectedGroup.Name, MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void Send()
    {
        if (SelectedGroup == null) return;
        var g = SelectedGroup;

        if (!HasAnyToRecipient(g))
        {
            MessageBox.Show("Add at least one recipient in the To field (uncheck CC on at least one row).",
                "No recipients",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var missing = g.Attachments.Where(a => !a.Exists).Select(a => a.Path).ToList();
        if (missing.Count > 0)
        {
            var list = string.Join(Environment.NewLine, missing);
            if (MessageBox.Show("These attachments are missing:\r\n" + list + "\r\n\r\nContinue without them?",
                    "Missing files", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;
        }

        var emptyFolders = g.Attachments
            .Where(a => a.IsFolder && !a.ResolveFiles().Any())
            .Select(a => a.Path).ToList();
        if (emptyFolders.Count > 0)
        {
            var list = string.Join(Environment.NewLine, emptyFolders);
            if (MessageBox.Show("These folder attachments contain no files:\r\n" + list + "\r\n\r\nSend anyway?",
                    "Empty folders", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;
        }

        var mode = _config.SendMode;

        if (mode == SendMode.SendAutomatically)
        {
            var confirm = MessageBox.Show(
                $"Send email from '{g.Name}' to {g.Recipients.Count} recipient(s) without review?",
                "Confirm auto-send", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;
        }

        try
        {
            SendGroupCore(g, mode);
            StatusMessage = mode == SendMode.SendAutomatically
                ? $"Sent '{g.Name}' to {g.Recipients.Count} recipient(s)."
                : $"Opened draft for '{g.Name}' in Outlook.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Outlook error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SendAll()
    {
        if (_config.Groups.Count == 0) return;

        var eligible = _config.Groups.Where(HasAnyToRecipient).ToList();
        var skipped  = _config.Groups.Except(eligible).ToList();

        if (eligible.Count == 0)
        {
            MessageBox.Show("No groups have a To recipient. Add at least one row (not CC-only) to a group.",
                "Nothing to send", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var missingByGroup = eligible
            .Select(g => (g, miss: g.Attachments.Where(a => !a.Exists).Select(a => a.Path).ToList()))
            .Where(x => x.miss.Count > 0)
            .ToList();

        var mode = _config.SendMode;
        var verb = mode == SendMode.SendAutomatically ? "Send" : "Open drafts for";
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"{verb} {eligible.Count} group(s):");
        foreach (var g in eligible)
            sb.AppendLine($"  • {g.Name}  ({g.Recipients.Count(r => !r.IsCc)} To, {g.Recipients.Count(r => r.IsCc)} Cc, {g.Attachments.Count} files)");
        if (skipped.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine($"Skipping {skipped.Count} group(s) with no To recipient:");
            foreach (var g in skipped) sb.AppendLine($"  • {g.Name}");
        }
        if (missingByGroup.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Some attachments are missing and will be omitted:");
            foreach (var (g, miss) in missingByGroup)
            {
                sb.AppendLine($"  • {g.Name}:");
                foreach (var p in miss) sb.AppendLine($"      - {p}");
            }
        }
        sb.AppendLine();
        sb.AppendLine("Continue?");

        var icon = mode == SendMode.SendAutomatically ? MessageBoxImage.Warning : MessageBoxImage.Question;
        if (MessageBox.Show(sb.ToString(), "Send all", MessageBoxButton.YesNo, icon) != MessageBoxResult.Yes)
            return;

        int ok = 0;
        var failures = new System.Collections.Generic.List<string>();
        foreach (var g in eligible)
        {
            try
            {
                SendGroupCore(g, mode);
                ok++;
            }
            catch (Exception ex)
            {
                failures.Add($"{g.Name}: {ex.Message}");
            }
        }

        StatusMessage = mode == SendMode.SendAutomatically
            ? $"Sent {ok} of {eligible.Count} group(s)."
            : $"Opened {ok} of {eligible.Count} draft(s) in Outlook.";

        if (failures.Count > 0)
        {
            MessageBox.Show(
                $"{failures.Count} group(s) failed:\r\n\r\n{string.Join(Environment.NewLine, failures)}",
                "Send all — errors", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static bool HasAnyToRecipient(EmailGroup g) =>
        g.Recipients.Any(r => !r.IsCc && !string.IsNullOrWhiteSpace(r.Email));

    /// <summary>
    /// Renders + sends/drafts a single group. No UI, no validation — caller must check first.
    /// Throws on Outlook errors.
    /// </summary>
    private void SendGroupCore(EmailGroup g, SendMode mode)
    {
        var subject = TemplateEngine.Render(g.Subject, g, _config);
        var body = TemplateEngine.Render(g.Body, g, _config);
        var toEmails = g.Recipients.Where(r => !r.IsCc && !string.IsNullOrWhiteSpace(r.Email)).Select(r => r.Email);
        var ccEmails = g.Recipients.Where(r =>  r.IsCc && !string.IsNullOrWhiteSpace(r.Email)).Select(r => r.Email);
        var paths = g.Attachments
            .Where(a => a.Exists)
            .SelectMany(a => a.ResolveFiles())
            .Distinct(StringComparer.OrdinalIgnoreCase);

        OutlookService.CreateMail(
            toEmails, ccEmails, subject, body, paths, _config.HtmlBody,
            mode == SendMode.SendAutomatically ? OutlookService.Mode.Send : OutlookService.Mode.Draft);

        CaptureContactsFrom(g);
    }
}
