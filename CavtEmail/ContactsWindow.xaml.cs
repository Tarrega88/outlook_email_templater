using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using CavtEmail.Models;

namespace CavtEmail;

public partial class ContactsWindow : Window
{
    private readonly ObservableCollection<Contact> _contacts;
    private readonly ICollectionView _view;

    public ContactsWindow(ObservableCollection<Contact> contacts)
    {
        InitializeComponent();
        _contacts = contacts;
        _view = CollectionViewSource.GetDefaultView(_contacts);
        _view.SortDescriptions.Add(new SortDescription(nameof(Contact.Name), ListSortDirection.Ascending));
        ContactList.DataContext = _view;
    }

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        var q = SearchBox.Text?.Trim() ?? "";
        if (string.IsNullOrEmpty(q))
        {
            _view.Filter = null;
        }
        else
        {
            _view.Filter = o =>
            {
                if (o is not Contact c) return false;
                return (c.Name?.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0)
                    || (c.Email?.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0);
            };
        }
    }

    private void Add_Click(object sender, RoutedEventArgs e)
    {
        var c = new Contact { Name = "", Email = "" };
        _contacts.Add(c);
        ContactList.ScrollIntoView(c);
        ContactList.SelectedItem = c;
    }

    private void Remove_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button b && b.DataContext is Contact c)
        {
            _contacts.Remove(c);
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
