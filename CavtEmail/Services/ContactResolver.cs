using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using CavtEmail.Models;

namespace CavtEmail.Services;

/// <summary>
/// Outcome of resolving a typed recipient string.
/// </summary>
public enum ResolveOutcome
{
    /// <summary>Exactly one match — <see cref="RecipientResolution.Contact"/> is set.</summary>
    Resolved,
    /// <summary>Multiple candidates — user must pick from <see cref="RecipientResolution.Candidates"/>.</summary>
    Ambiguous,
    /// <summary>No match anywhere and text is not a valid email literal.</summary>
    Unresolved,
    /// <summary>No match but input is a plausible email; accepted as a literal external address.</summary>
    AcceptedLiteral
}

public sealed class RecipientResolution
{
    public ResolveOutcome Outcome { get; init; }
    public Contact? Contact { get; init; }
    public IReadOnlyList<Contact> Candidates { get; init; } = Array.Empty<Contact>();
}

/// <summary>
/// Resolves a free-form query string (name / alias / email) to a <see cref="Contact"/>.
///
/// Strategy:
///   1. Exact email match against local contacts
///   2. Exact alias match
///   3. Prefix match on Name or Aliases (single match wins; many -> Ambiguous)
///   4. Outlook GAL resolve (single call, bounded time)
///   5. If query looks like an email, accept it as a literal external address
///   6. Otherwise Unresolved
///
/// Successful Outlook hits are added to the local book with the raw query
/// stored as an alias, so future lookups for the same string are instant
/// and work offline.
/// </summary>
public class ContactResolver
{
    private static readonly Regex EmailRegex = new(
        @"^[^\s@]+@[^\s@]+\.[^\s@]+$", RegexOptions.Compiled);

    private readonly System.Collections.ObjectModel.ObservableCollection<Contact> _contacts;
    private readonly Action _onChanged; // called when the book is mutated so caller can persist

    /// <summary>
    /// Emails (lowercased) we've already asked Outlook to verify this session.
    /// Prevents re-querying for contacts Outlook doesn't know (e.g. external vendors).
    /// </summary>
    private readonly HashSet<string> _outlookUpgradeTried = new(StringComparer.OrdinalIgnoreCase);

    public ContactResolver(
        System.Collections.ObjectModel.ObservableCollection<Contact> contacts,
        Action onChanged)
    {
        _contacts = contacts;
        _onChanged = onChanged;
    }

    /// <summary>True if classic Outlook is currently running (process check, no COM call).</summary>
    public static bool IsOutlookRunning()
    {
        try { return Process.GetProcessesByName("OUTLOOK").Length > 0; }
        catch { return false; }
    }

    public static bool LooksLikeEmail(string? s) =>
        !string.IsNullOrWhiteSpace(s) && EmailRegex.IsMatch(s.Trim());

    public RecipientResolution Resolve(string? rawQuery)
    {
        var q = (rawQuery ?? "").Trim();
        if (q.Length == 0)
            return new RecipientResolution { Outcome = ResolveOutcome.Unresolved };

        // 1. Exact email match.
        var byEmail = _contacts.FirstOrDefault(c =>
            !string.IsNullOrWhiteSpace(c.Email) &&
            string.Equals(c.Email, q, StringComparison.OrdinalIgnoreCase));
        if (byEmail != null)
        {
            TryUpgradeFromOutlook(byEmail);
            Stamp(byEmail, q);
            return Ok(byEmail);
        }

        // 2. Exact alias or name match.
        var exact = _contacts.Where(c =>
            string.Equals(c.Name, q, StringComparison.OrdinalIgnoreCase) ||
            c.Aliases.Any(a => string.Equals(a, q, StringComparison.OrdinalIgnoreCase))
        ).ToList();
        if (exact.Count == 1) { TryUpgradeFromOutlook(exact[0]); Stamp(exact[0], q); return Ok(exact[0]); }
        if (exact.Count > 1)  return Ambiguous(exact);

        // 3. Prefix match on name or any alias.
        var prefix = _contacts.Where(c =>
            (c.Name?.StartsWith(q, StringComparison.OrdinalIgnoreCase) ?? false) ||
            c.Aliases.Any(a => a.StartsWith(q, StringComparison.OrdinalIgnoreCase))
        ).ToList();
        if (prefix.Count == 1) { TryUpgradeFromOutlook(prefix[0]); AddAlias(prefix[0], q); Stamp(prefix[0], q); return Ok(prefix[0]); }
        if (prefix.Count > 1)  return Ambiguous(prefix);

        // 4. Ask Outlook (only if it's running — we promised not to launch it).
        if (IsOutlookRunning())
        {
            var gal = TryOutlookResolve(q);
            if (gal.Outcome == ResolveOutcome.Resolved && gal.Contact != null)
            {
                var c = UpsertFromOutlook(gal.Contact, q);
                return Ok(c);
            }
            if (gal.Outcome == ResolveOutcome.Ambiguous)
                return gal;
        }

        // 5. Accept literal emails (e.g. external vendors) but DO NOT persist them
        //    yet — otherwise a typo gets memorialized in the cache forever.
        //    The caller shows an "unverified" warning state instead.
        if (LooksLikeEmail(q))
        {
            return new RecipientResolution
            {
                Outcome = ResolveOutcome.AcceptedLiteral,
                Contact = new Contact { Name = "", Email = q }
            };
        }

        return new RecipientResolution { Outcome = ResolveOutcome.Unresolved };
    }

    /// <summary>Commit a user-selected candidate from an ambiguous result.</summary>
    public void ConfirmPick(Contact picked, string originalQuery)
    {
        AddAlias(picked, originalQuery);
        Stamp(picked, originalQuery);
    }

    /// <summary>
    /// Cache-only suggestions for live autocomplete. Fast path — no Outlook
    /// calls. Matches are scored:
    ///   exact email / name / alias (case-insensitive) > prefix > substring.
    /// Within a tier, most-recently-used wins.
    /// </summary>
    public IReadOnlyList<Contact> Suggest(string query, int limit = 8)
    {
        var q = (query ?? "").Trim();
        if (q.Length < 2) return Array.Empty<Contact>();

        int Score(Contact c)
        {
            bool EqAny(string? s) => s != null && string.Equals(s, q, StringComparison.OrdinalIgnoreCase);
            bool StartsAny(string? s) => s != null && s.StartsWith(q, StringComparison.OrdinalIgnoreCase);
            bool ContainsAny(string? s) => s != null && s.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0;

            if (EqAny(c.Email) || EqAny(c.Name) || c.Aliases.Any(a => EqAny(a))) return 3;
            if (StartsAny(c.Name) || StartsAny(c.Email) || c.Aliases.Any(a => StartsAny(a))) return 2;
            if (ContainsAny(c.Name) || ContainsAny(c.Email) || c.Aliases.Any(a => ContainsAny(a))) return 1;
            return 0;
        }

        return _contacts
            .Select(c => (c, s: Score(c)))
            .Where(t => t.s > 0)
            .OrderByDescending(t => t.s)
            .ThenByDescending(t => t.c.LastUsedUtc)
            .Take(limit)
            .Select(t => t.c)
            .ToList();
    }

    /// <summary>
    /// Persist a user-confirmed "unverified" address to the local cache with the
    /// name they supplied. Marked NOT FoundInOutlook — the green ● icon is correct.
    /// Returns the stored contact.
    /// </summary>
    public Contact SaveManual(string name, string email)
    {
        var existing = _contacts.FirstOrDefault(c =>
            string.Equals(c.Email, email, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            if (!string.IsNullOrWhiteSpace(name)) existing.Name = name;
            Stamp(existing, email);
            return existing;
        }
        return Upsert(new Contact { Name = name ?? "", Email = email ?? "" }, email ?? "");
    }

    private RecipientResolution Ok(Contact c) =>
        new() { Outcome = ResolveOutcome.Resolved, Contact = c };

    private RecipientResolution Ambiguous(List<Contact> list) =>
        new()
        {
            Outcome = ResolveOutcome.Ambiguous,
            Candidates = list.OrderByDescending(c => c.LastUsedUtc).ToList()
        };

    // --------------------------------------------------------------
    // Mutations
    // --------------------------------------------------------------

    private Contact UpsertFromOutlook(Contact fromOutlook, string originalQuery)
    {
        var existing = _contacts.FirstOrDefault(c =>
            string.Equals(c.Email, fromOutlook.Email, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            // Refresh name if Outlook knows better.
            if (!string.IsNullOrWhiteSpace(fromOutlook.Name) && existing.Name != fromOutlook.Name)
                existing.Name = fromOutlook.Name;
            // Upgrade trust level — once Outlook confirms it, it stays Outlook-confirmed.
            if (!existing.FoundInOutlook) existing.FoundInOutlook = true;
            AddAlias(existing, originalQuery);
            Stamp(existing, originalQuery);
            return existing;
        }
        fromOutlook.FoundInOutlook = true;
        return Upsert(fromOutlook, originalQuery);
    }

    private Contact Upsert(Contact incoming, string originalQuery)
    {
        var c = new Contact
        {
            Name = incoming.Name ?? "",
            Email = incoming.Email ?? "",
            LastUsedUtc = DateTime.UtcNow,
            FoundInOutlook = incoming.FoundInOutlook
        };
        SeedAliases(c);
        AddAlias(c, originalQuery);
        _contacts.Add(c);
        _onChanged();
        return c;
    }

    private void SeedAliases(Contact c)
    {
        if (!string.IsNullOrWhiteSpace(c.Name)) AddAliasSilent(c, c.Name);
        var first = NameUtil.FirstName(c.Name);
        if (!string.IsNullOrEmpty(first)) AddAliasSilent(c, first!);
        var last = NameUtil.LastName(c.Name);
        if (!string.IsNullOrEmpty(last))  AddAliasSilent(c, last!);
        if (!string.IsNullOrWhiteSpace(c.Email))
        {
            var at = c.Email.IndexOf('@');
            if (at > 0) AddAliasSilent(c, c.Email[..at]);
        }
    }

    private void AddAlias(Contact c, string alias)
    {
        if (AddAliasSilent(c, alias)) _onChanged();
    }

    private static bool AddAliasSilent(Contact c, string alias)
    {
        alias = alias?.Trim() ?? "";
        if (alias.Length == 0) return false;
        if (c.Aliases.Any(a => string.Equals(a, alias, StringComparison.OrdinalIgnoreCase))) return false;
        c.Aliases.Add(alias);
        return true;
    }

    private void Stamp(Contact c, string _)
    {
        c.LastUsedUtc = DateTime.UtcNow;
        _onChanged();
    }

    /// <summary>
    /// If <paramref name="c"/> is a cache-only contact (added before the FoundInOutlook
    /// field existed, or saved when Outlook was offline), try once per session to
    /// verify it against Outlook so the origin icon upgrades from green to blue.
    /// Cheap: Outlook.Resolve on a literal email is a single RPC, and we skip entirely
    /// if Outlook isn't running or we've already asked about this email this session.
    /// </summary>
    private void TryUpgradeFromOutlook(Contact c)
    {
        if (c.FoundInOutlook) return;
        if (string.IsNullOrWhiteSpace(c.Email)) return;
        if (!_outlookUpgradeTried.Add(c.Email)) return;
        if (!IsOutlookRunning()) return;

        var gal = TryOutlookResolve(c.Email);
        if (gal.Outcome == ResolveOutcome.Resolved && gal.Contact != null)
        {
            c.FoundInOutlook = true;
            if (!string.IsNullOrWhiteSpace(gal.Contact.Name) && c.Name != gal.Contact.Name)
                c.Name = gal.Contact.Name;
            _onChanged();
        }
    }

    // --------------------------------------------------------------
    // Outlook COM bridge
    // --------------------------------------------------------------

    /// <summary>
    /// Ask Outlook to resolve <paramref name="q"/> against its address book(s).
    /// Uses late-bound COM so no Outlook PIA is required.
    /// Returns <c>Unresolved</c> on any failure — callers fall back gracefully.
    /// </summary>
    private static RecipientResolution TryOutlookResolve(string q)
    {
        dynamic? outlook = null;
        dynamic? session = null;
        dynamic? recip = null;
        try
        {
            var appType = Type.GetTypeFromProgID("Outlook.Application");
            if (appType == null)
                return new RecipientResolution { Outcome = ResolveOutcome.Unresolved };

            outlook = Activator.CreateInstance(appType);
            if (outlook == null)
                return new RecipientResolution { Outcome = ResolveOutcome.Unresolved };

            session = outlook.Session;
            recip = session.CreateRecipient(q);
            bool ok = recip.Resolve();
            if (!ok)
                return new RecipientResolution { Outcome = ResolveOutcome.Unresolved };

            var c = ContactFromRecipient(recip);
            return c == null
                ? new RecipientResolution { Outcome = ResolveOutcome.Unresolved }
                : new RecipientResolution { Outcome = ResolveOutcome.Resolved, Contact = c };
        }
        catch
        {
            return new RecipientResolution { Outcome = ResolveOutcome.Unresolved };
        }
        finally
        {
            if (recip   != null) Marshal.FinalReleaseComObject(recip);
            if (session != null) Marshal.FinalReleaseComObject(session);
            if (outlook != null) Marshal.FinalReleaseComObject(outlook);
        }
    }

    private static Contact? ContactFromRecipient(dynamic recip)
    {
        try
        {
            string name = recip.Name as string ?? "";
            string? email = null;
            bool isRealDirectoryEntry = false;

            // Preferred: Exchange user's primary SMTP.
            try
            {
                var entry = recip.AddressEntry;
                if (entry != null)
                {
                    // AddressEntryUserType tells us if this is a real address book entry
                    // or a one-off SMTP literal that Outlook just rubber-stamped.
                    //
                    //    0  = olExchangeUserAddressEntry
                    //    1  = olExchangeDistributionListAddressEntry
                    //    2  = olExchangePublicFolderAddressEntry
                    //    3  = olExchangeAgentAddressEntry
                    //    4  = olExchangeOrganizationAddressEntry
                    //    5  = olExchangeRemoteUserAddressEntry
                    //   10  = olOutlookContactAddressEntry
                    //   11  = olOutlookDistributionListAddressEntry
                    //   20  = olLdapAddressEntry
                    //   30  = olSmtpAddressEntry          <-- one-off / literal
                    //   40  = olOtherAddressEntry
                    //
                    // Anything other than the SMTP / Other buckets is a real hit.
                    int userType = 30;
                    try { userType = (int)entry.AddressEntryUserType; } catch { }
                    isRealDirectoryEntry = userType != 30 && userType != 40;

                    try
                    {
                        var xu = entry.GetExchangeUser();
                        if (xu != null)
                        {
                            email = xu.PrimarySmtpAddress as string;
                            if (string.IsNullOrWhiteSpace(name))
                                name = (xu.Name as string) ?? name;
                        }
                    }
                    catch { /* not an Exchange user */ }

                    if (string.IsNullOrWhiteSpace(email))
                    {
                        // SMTP address if one-off or contact
                        try { email = entry.Address as string; } catch { }
                    }
                }
            }
            catch { }

            if (string.IsNullOrWhiteSpace(email)) return null;
            if (!isRealDirectoryEntry) return null; // let caller fall through to AcceptedLiteral
            return new Contact { Name = name?.Trim() ?? "", Email = email!.Trim() };
        }
        catch
        {
            return null;
        }
    }
}
