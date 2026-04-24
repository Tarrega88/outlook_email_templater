using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using CavtEmail.Models;

namespace CavtEmail.Services;

/// <summary>
/// Contacts are stored in their own file (shared across every config the user
/// opens/saves), so adding a person once makes them available everywhere.
/// </summary>
public static class ContactsService
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    public static string DefaultPath
    {
        get
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CavtEmail");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "contacts.json");
        }
    }

    public static ObservableCollection<Contact> Load()
    {
        var path = DefaultPath;
        if (!File.Exists(path)) return new ObservableCollection<Contact>();
        try
        {
            var json = File.ReadAllText(path);
            var list = JsonSerializer.Deserialize<List<Contact>>(json, Options)
                       ?? new List<Contact>();
            return new ObservableCollection<Contact>(list);
        }
        catch
        {
            return new ObservableCollection<Contact>();
        }
    }

    public static void Save(IEnumerable<Contact> contacts)
    {
        var path = DefaultPath;
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        var json = JsonSerializer.Serialize(contacts.ToList(), Options);
        File.WriteAllText(path, json);
    }

    /// <summary>
    /// Merge a set of contacts into the target collection, deduping by email
    /// (case-insensitive). Returns the number of new contacts added.
    /// </summary>
    public static int Merge(ObservableCollection<Contact> target, IEnumerable<Contact> incoming)
    {
        int added = 0;
        foreach (var c in incoming)
        {
            if (string.IsNullOrWhiteSpace(c.Email)) continue;
            if (target.Any(x => string.Equals(x.Email, c.Email, StringComparison.OrdinalIgnoreCase)))
                continue;
            target.Add(new Contact
            {
                Name = c.Name,
                Email = c.Email,
                FoundInOutlook = c.FoundInOutlook,
                LastUsedUtc = c.LastUsedUtc,
                Aliases = c.Aliases != null ? new List<string>(c.Aliases) : new List<string>()
            });
            added++;
        }
        return added;
    }
}
