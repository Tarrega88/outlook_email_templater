using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Envoy.Models;
using Envoy.Services;
using Envoy.ViewModels;

namespace Envoy;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private MainViewModel VM => (MainViewModel)DataContext;

    private void ExitMenu_Click(object sender, RoutedEventArgs e) => Close();

    // ------------------------------------------------------------------
    // Attachments drag-and-drop
    // ------------------------------------------------------------------

    private void Attachments_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    private void Attachments_Drop(object sender, DragEventArgs e)
    {
        if (VM.SelectedEmail == null) return;
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        if (e.Data.GetData(DataFormats.FileDrop) is not string[] items) return;

        foreach (var p in items)
        {
            // Folders are kept as references (resolved at send time); files are added directly.
            if (Directory.Exists(p) || File.Exists(p))
                VM.AddAttachmentPath(p);
        }
        e.Handled = true;
    }

    // ------------------------------------------------------------------
    // Token / row insertion
    // ------------------------------------------------------------------

    /// <summary>Where a row's Subj/Body button inserts text.</summary>
    private enum Target { Body, Subject }

    private static Target ParseTarget(object? tag) =>
        string.Equals(tag as string, "Subject", StringComparison.OrdinalIgnoreCase)
            ? Target.Subject
            : Target.Body;

    private void InsertAtCursor(TextBox tb, string text)
    {
        if (string.IsNullOrEmpty(text)) return;

        var caret = tb.CaretIndex;
        if (tb.SelectionLength > 0)
        {
            tb.SelectedText = text;
            tb.CaretIndex = caret + text.Length;
        }
        else
        {
            tb.Text = (tb.Text ?? "").Insert(caret, text);
            tb.CaretIndex = caret + text.Length;
        }
        tb.Focus();
    }

    private void Insert(Target target, string text)
    {
        var tb = target == Target.Subject ? SubjectBox : BodyBox;
        InsertAtCursor(tb, text);
    }

    private static Target TargetFromSender(object sender)
    {
        if (sender is FrameworkElement fe)
            return ParseTarget(fe.Tag);
        return Target.Body;
    }

    private void InsertAttachment_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.DataContext is AttachmentItem a)
        {
            Insert(TargetFromSender(sender), a.FileName);
        }
    }

    /// <summary>Token pills on the right panel; honors the most-recently-focused box.</summary>
    private void Token_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button b && b.Content is string s)
        {
            Insert(_lastTokenTarget, s);
        }
    }

    private Target _lastTokenTarget = Target.Body;

    private void SubjectBody_GotFocus(object sender, RoutedEventArgs e)
    {
        if (ReferenceEquals(sender, SubjectBox))
        {
            _lastTokenTarget = Target.Subject;
            if (TokenTargetLabel != null) TokenTargetLabel.Text = "Subject";
        }
        else if (ReferenceEquals(sender, BodyBox))
        {
            _lastTokenTarget = Target.Body;
            if (TokenTargetLabel != null) TokenTargetLabel.Text = "Body";
        }
    }

    /// <summary>Right-click context menu on Subject/Body — inserts the menu item's Tag at the caret of the box that opened the menu.</summary>
    private void TokenMenu_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem mi || mi.Tag is not string token) return;

        // Walk up to the owning ContextMenu, then to its PlacementTarget (the TextBox).
        var parent = mi.Parent;
        while (parent is MenuItem outer) parent = outer.Parent;
        if (parent is ContextMenu cm && cm.PlacementTarget is TextBox tb)
        {
            InsertAtCursor(tb, token);
        }
        else
        {
            // Fallback: insert into the most recently focused box.
            Insert(_lastTokenTarget, token);
        }
    }

    // ------------------------------------------------------------------
    // Recipient resolve-on-commit
    // ------------------------------------------------------------------

    private Recipient? _ambiguousFor;

    private void RecipientBox_GotFocus(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox tb)
        {
            // Select-all so retyping replaces the previous query cleanly.
            tb.Dispatcher.BeginInvoke(new Action(() => tb.SelectAll()),
                System.Windows.Threading.DispatcherPriority.Input);
        }
    }

    private void RecipientBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox tb && tb.DataContext is Recipient r)
        {
            // Give the autocomplete popup a moment to receive its own click; only
            // resolve if focus genuinely moved away from both.
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_autocompleteJustAccepted) { _autocompleteJustAccepted = false; return; }
                if (Autocomplete.IsOpen &&
                    (Autocomplete.IsMouseOver || AutocompleteList.IsMouseOver ||
                     AutocompleteList.IsKeyboardFocusWithin))
                    return;
                CloseAutocomplete();
                ResolveRecipient(r);
            }), DispatcherPriority.Background);
        }
    }

    private void RecipientBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (sender is not TextBox tb || tb.DataContext is not Recipient r) return;

        // Route navigation keys into the autocomplete popup when it's open.
        if (Autocomplete.IsOpen)
        {
            if (e.Key == Key.Down)
            {
                AutocompleteList.SelectedIndex = Math.Min(
                    AutocompleteList.Items.Count - 1,
                    Math.Max(0, AutocompleteList.SelectedIndex + 1));
                ScrollAutocompleteIntoView();
                e.Handled = true;
                return;
            }
            if (e.Key == Key.Up)
            {
                AutocompleteList.SelectedIndex = Math.Max(0, AutocompleteList.SelectedIndex - 1);
                ScrollAutocompleteIntoView();
                e.Handled = true;
                return;
            }
            if ((e.Key == Key.Enter || e.Key == Key.Tab) &&
                AutocompleteList.SelectedItem is Contact picked)
            {
                AcceptAutocomplete(r, picked);
                e.Handled = true;
                return;
            }
            if (e.Key == Key.Escape)
            {
                CloseAutocomplete();
                _autocompleteDismissed = true;
                e.Handled = true;
                return;
            }
        }

        if (e.Key == Key.Enter)
        {
            ResolveRecipient(r);
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            // Revert the in-progress query to whatever was already committed.
            r.Query = string.IsNullOrWhiteSpace(r.Name) ? r.Email : r.Name;
            Keyboard.ClearFocus();
            e.Handled = true;
        }
    }

    // --------------------------------------------------------------
    // Autocomplete (cache-only, 300ms debounce)
    // --------------------------------------------------------------

    private readonly DispatcherTimer _autocompleteTimer = new() { Interval = TimeSpan.FromMilliseconds(300) };
    private TextBox? _autocompleteFor;
    private bool _autocompleteDismissed;
    private bool _autocompleteTimerWired;
    private bool _autocompleteJustAccepted;

    private void EnsureAutocompleteWired()
    {
        if (_autocompleteTimerWired) return;
        _autocompleteTimerWired = true;
        _autocompleteTimer.Tick += (_, _) =>
        {
            _autocompleteTimer.Stop();
            RunAutocomplete();
        };
    }

    private void RecipientBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is not TextBox tb || tb.DataContext is not Recipient r) return;
        EnsureAutocompleteWired();
        _autocompleteFor = tb;

        // User typed — any previous dismissal no longer applies.
        _autocompleteDismissed = false;

        // Don't autocomplete while the field is showing something already resolved.
        if (r.IsResolved || r.IsResolving)
        {
            CloseAutocomplete();
            return;
        }

        _autocompleteTimer.Stop();
        _autocompleteTimer.Start();
    }

    private void RunAutocomplete()
    {
        if (_autocompleteFor == null || _autocompleteDismissed) return;
        if (_autocompleteFor.DataContext is not Recipient r) return;
        if (!_autocompleteFor.IsKeyboardFocused) { CloseAutocomplete(); return; }

        var q = (r.Query ?? "").Trim();
        var used = CurrentRecipientEmails(r);
        var hits = VM.Resolver.Suggest(q)
            .Where(c => !string.IsNullOrWhiteSpace(c.Email) && !used.Contains(c.Email!))
            .ToList();
        if (hits.Count == 0)
        {
            CloseAutocomplete();
            return;
        }

        AutocompleteList.ItemsSource = hits;
        AutocompleteList.SelectedIndex = 0;
        Autocomplete.PlacementTarget = _autocompleteFor;
        Autocomplete.Width = _autocompleteFor.ActualWidth;
        Autocomplete.IsOpen = true;
    }

    /// <summary>Emails already used by other recipients on the current email (case-insensitive).</summary>
    private HashSet<string> CurrentRecipientEmails(Recipient exclude)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (VM.SelectedEmail is null) return set;
        foreach (var other in VM.SelectedEmail.Recipients)
        {
            if (ReferenceEquals(other, exclude)) continue;
            if (!string.IsNullOrWhiteSpace(other.Email)) set.Add(other.Email!);
        }
        return set;
    }

    private void CloseAutocomplete()
    {
        _autocompleteTimer.Stop();
        Autocomplete.IsOpen = false;
    }

    private void ScrollAutocompleteIntoView()
    {
        if (AutocompleteList.SelectedItem != null)
            AutocompleteList.ScrollIntoView(AutocompleteList.SelectedItem);
    }

    private void AutocompleteList_Click(object sender, MouseButtonEventArgs e)
    {
        // Walk up from the click source to find the row (ListBoxItem).
        var hit = e.OriginalSource as DependencyObject;
        while (hit != null && hit is not ListBoxItem)
            hit = System.Windows.Media.VisualTreeHelper.GetParent(hit);
        if (hit is ListBoxItem lbi && lbi.DataContext is Contact c &&
            _autocompleteFor?.DataContext is Recipient r)
        {
            AcceptAutocomplete(r, c);
            e.Handled = true;
        }
    }

    private void AutocompleteList_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape) { CloseAutocomplete(); _autocompleteDismissed = true; e.Handled = true; }
    }

    private void AcceptAutocomplete(Recipient r, Contact c)
    {
        if (IsDuplicate(r, c.Email))
        {
            RevertRecipient(r);
            CloseAutocomplete();
            return;
        }
        VM.Resolver.ConfirmPick(c, (r.Query ?? "").Trim());
        r.Name = c.Name ?? "";
        r.Email = c.Email ?? "";
        r.Query = string.IsNullOrWhiteSpace(r.Name) ? r.Email : r.Name;
        r.State = c.FoundInOutlook ? ResolveState.ResolvedOutlook : ResolveState.ResolvedLocal;
        VM.EnsureTrailingBlankRecipient(VM.SelectedEmail);
        _autocompleteJustAccepted = true;
        CloseAutocomplete();
    }

    /// <summary>True when <paramref name="email"/> is already used by another recipient on the current email.</summary>
    private bool IsDuplicate(Recipient r, string? email)
    {
        if (string.IsNullOrWhiteSpace(email) || VM.SelectedEmail is null) return false;
        foreach (var other in VM.SelectedEmail.Recipients)
        {
            if (ReferenceEquals(other, r)) continue;
            if (string.Equals(other.Email, email, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private void WarnDuplicate(string? email) { /* intentionally silent per UX */ }

    private static void RevertRecipient(Recipient r)
    {
        r.Name = "";
        r.Email = "";
        r.Query = "";
        r.State = ResolveState.Empty;
    }

    private void EditRecipient_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.DataContext is not Recipient r) return;

        // Demote to Empty so the TextBox becomes visible again. We seed Query
        // from Name so the user sees what they had.
        r.Query = string.IsNullOrWhiteSpace(r.Name) ? r.Email : r.Name;
        r.State = ResolveState.Empty;

        // After the template swap, find and focus the TextBox that corresponds to this row.
        Dispatcher.BeginInvoke(new Action(() =>
        {
            if (FindTextBoxForRecipient(btn) is TextBox tb)
            {
                tb.Focus();
                tb.SelectAll();
            }
        }), System.Windows.Threading.DispatcherPriority.Input);
    }

    /// <summary>Walks up from the pill button to find the sibling TextBox in the row template.</summary>
    private static TextBox? FindTextBoxForRecipient(DependencyObject start)
    {
        var parent = System.Windows.Media.VisualTreeHelper.GetParent(start);
        while (parent != null && parent is not Grid) parent = System.Windows.Media.VisualTreeHelper.GetParent(parent);
        if (parent is not Grid g) return null;
        for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(g); i++)
        {
            if (System.Windows.Media.VisualTreeHelper.GetChild(g, i) is TextBox tb) return tb;
        }
        return null;
    }

    private void ResolveRecipient(Recipient r)
    {
        var q = (r.Query ?? "").Trim();

        if (q.Length == 0)
        {
            // Blank field — treat as empty, clear stored name/email so Send won't pick up stale data.
            r.Name = "";
            r.Email = "";
            r.State = ResolveState.Empty;
            return;
        }

        // If nothing changed vs. the committed state, don't bother re-resolving.
        if (r.IsResolved &&
            (string.Equals(q, r.Name, StringComparison.OrdinalIgnoreCase) ||
             string.Equals(q, r.Email, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        r.State = ResolveState.Resolving;

        var result = VM.Resolver.Resolve(q);
        switch (result.Outcome)
        {
            case ResolveOutcome.Resolved:
            {
                var c = result.Contact!;
                if (IsDuplicate(r, c.Email)) { WarnDuplicate(c.Email); RevertRecipient(r); return; }
                r.Name = c.Name ?? "";
                r.Email = c.Email ?? "";
                r.Query = string.IsNullOrWhiteSpace(r.Name) ? r.Email : r.Name;
                r.State = c.FoundInOutlook ? ResolveState.ResolvedOutlook : ResolveState.ResolvedLocal;
                VM.EnsureTrailingBlankRecipient(VM.SelectedEmail);
                break;
            }

            case ResolveOutcome.AcceptedLiteral:
            {
                // Valid email format but unknown to Outlook/cache. Flag it, don't persist.
                var c = result.Contact!;
                if (IsDuplicate(r, c.Email)) { WarnDuplicate(c.Email); RevertRecipient(r); return; }
                r.Name = c.Name ?? "";
                r.Email = c.Email ?? "";
                r.Query = r.Email;
                r.State = ResolveState.Unverified;
                VM.EnsureTrailingBlankRecipient(VM.SelectedEmail);
                break;
            }

            case ResolveOutcome.Ambiguous:
                ShowAmbiguous(r, q, result.Candidates);
                break;

            case ResolveOutcome.Unresolved:
            default:
                r.State = ResolveState.Unresolved;
                break;
        }
    }

    // ------------------------------------------------------------------
    // Ambiguous-candidates popup
    // ------------------------------------------------------------------

    private string _ambiguousQuery = "";

    private void ShowAmbiguous(Recipient r, string query, IReadOnlyList<Contact> candidates)
    {
        _ambiguousFor = r;
        _ambiguousQuery = query;
        AmbiguousList.ItemsSource = candidates;
        AmbiguousList.SelectedIndex = 0;
        AmbiguousPicker.PlacementTarget = FindTextBoxForRecipient(AmbiguousPicker) ?? (UIElement)this;
        // Better: anchor to the focused TextBox if any.
        if (Keyboard.FocusedElement is TextBox focused) AmbiguousPicker.PlacementTarget = focused;
        AmbiguousPicker.IsOpen = true;
        r.State = ResolveState.Unresolved; // until user picks
    }

    private void AmbiguousCommit_Click(object sender, MouseButtonEventArgs e) => CommitAmbiguous();

    private void AmbiguousList_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) { CommitAmbiguous(); e.Handled = true; }
        else if (e.Key == Key.Escape) { AmbiguousPicker.IsOpen = false; e.Handled = true; }
    }

    private void CommitAmbiguous()
    {
        if (_ambiguousFor == null || AmbiguousList.SelectedItem is not Contact c)
        {
            AmbiguousPicker.IsOpen = false;
            return;
        }
        if (IsDuplicate(_ambiguousFor, c.Email))
        {
            WarnDuplicate(c.Email);
            RevertRecipient(_ambiguousFor);
            AmbiguousPicker.IsOpen = false;
            _ambiguousFor = null;
            return;
        }
        VM.Resolver.ConfirmPick(c, _ambiguousQuery);
        _ambiguousFor.Name = c.Name ?? "";
        _ambiguousFor.Email = c.Email ?? "";
        _ambiguousFor.Query = string.IsNullOrWhiteSpace(_ambiguousFor.Name) ? _ambiguousFor.Email : _ambiguousFor.Name;
        _ambiguousFor.State = c.FoundInOutlook ? ResolveState.ResolvedOutlook : ResolveState.ResolvedLocal;
        VM.EnsureTrailingBlankRecipient(VM.SelectedEmail);
        AmbiguousPicker.IsOpen = false;
        _ambiguousFor = null;
    }

    // ------------------------------------------------------------------
    // Status-icon click (amber "?" only -> add-to-contacts dialog)
    // ------------------------------------------------------------------

    private void StatusIcon_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.DataContext is not Recipient r) return;
        if (!r.IsUnverified || string.IsNullOrWhiteSpace(r.Email)) return;

        var dlg = new AddContactDialog(r.Email) { Owner = this };
        if (dlg.ShowDialog() != true) return;

        var saved = VM.Resolver.SaveManual(dlg.FullName, r.Email);
        r.Name = saved.Name ?? "";
        r.Email = saved.Email ?? "";
        r.Query = string.IsNullOrWhiteSpace(r.Name) ? r.Email : r.Name;
        r.State = saved.FoundInOutlook ? ResolveState.ResolvedOutlook : ResolveState.ResolvedLocal;
    }
}