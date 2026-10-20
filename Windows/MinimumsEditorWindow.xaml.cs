using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using HexxcorpDocumentReviewUtility.Models;
using HexxcorpDocumentReviewUtility.Services;

namespace HexxcorpDocumentReviewUtility.Windows
{
    /// <summary>
    /// Simple editor for the DOT starter and user-configured criteria:
    /// add/remove/reorder sections, and edit each section's title and lines.
    /// Saves back through <see cref="MinimumsDataService"/>; the owning
    /// <see cref="MinimumsWindow"/> reloads its display after a successful save.
    /// </summary>
    public partial class MinimumsEditorWindow : Window
    {
        private readonly ObservableCollection<MinimumsSection> _sections;
        private bool _isLoadingSelection;

        public MinimumsEditorWindow()
        {
            InitializeComponent();

            MinimumsData data = MinimumsDataService.Load();
            _sections = new ObservableCollection<MinimumsSection>(data.Sections);
            SectionsListBox.ItemsSource = _sections;

            if (_sections.Count > 0)
            {
                SectionsListBox.SelectedIndex = 0;
            }
            else
            {
                SetEditorEnabled(false);
            }
        }

        private MinimumsSection? SelectedSection => SectionsListBox.SelectedItem as MinimumsSection;

        private void SetEditorEnabled(bool enabled)
        {
            TitleTextBox.IsEnabled = enabled;
            LinesTextBox.IsEnabled = enabled;
        }

        private void SectionsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _isLoadingSelection = true;
            try
            {
                MinimumsSection? section = SelectedSection;
                if (section is null)
                {
                    TitleTextBox.Text = string.Empty;
                    LinesTextBox.Text = string.Empty;
                    SetEditorEnabled(false);
                    return;
                }

                SetEditorEnabled(true);
                TitleTextBox.Text = section.Title;
                LinesTextBox.Text = string.Join(Environment.NewLine, section.Lines);
            }
            finally
            {
                _isLoadingSelection = false;
            }
        }

        private void TitleTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isLoadingSelection || SelectedSection is not { } section)
            {
                return;
            }

            section.Title = TitleTextBox.Text;
        }

        private void LinesTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isLoadingSelection || SelectedSection is not { } section)
            {
                return;
            }

            section.Lines = LinesTextBox.Text
                .Replace("\r\n", "\n")
                .Split('\n')
                .ToList();
        }

        private void AddSectionButton_Click(object sender, RoutedEventArgs e)
        {
            var section = new MinimumsSection
            {
                Title = "New Section",
                Lines = { "Enter details here." }
            };

            _sections.Add(section);
            SectionsListBox.SelectedItem = section;

            TitleTextBox.Focus();
            TitleTextBox.SelectAll();
        }

        private void RemoveSectionButton_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedSection is not { } section)
            {
                return;
            }

            var result = MessageBox.Show(
                this,
                $"Remove the section \"{section.Title}\"? This cannot be undone.",
                "Remove Section",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (result == MessageBoxResult.Yes)
            {
                int index = _sections.IndexOf(section);
                _sections.Remove(section);
                if (_sections.Count > 0)
                {
                    SectionsListBox.SelectedIndex = Math.Min(index, _sections.Count - 1);
                }
            }
        }

        private void MoveSectionUpButton_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedSection is not { } section)
            {
                return;
            }

            int index = _sections.IndexOf(section);
            if (index > 0)
            {
                _sections.Move(index, index - 1);
            }
        }

        private void MoveSectionDownButton_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedSection is not { } section)
            {
                return;
            }

            int index = _sections.IndexOf(section);
            if (index >= 0 && index < _sections.Count - 1)
            {
                _sections.Move(index, index + 1);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            var data = new MinimumsData { Sections = _sections.ToList() };
            bool saved = MinimumsDataService.Save(data);

            if (saved)
            {
                DialogResult = true;
            }
            else
            {
                StatusText.Text = "Could not save - the file may be locked or read-only.";
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
