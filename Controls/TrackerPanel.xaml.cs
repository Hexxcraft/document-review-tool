using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using HexxcorpDocumentReviewUtility.Models;

namespace HexxcorpDocumentReviewUtility.Controls
{
    /// <summary>
    /// Shared Dates &amp; Employment History tracker view. DataContext is a <see cref="TrackerModel"/>,
    /// so every instance (main drawer, workspace drawer) shows the same data.
    /// </summary>
    public partial class TrackerPanel : UserControl
    {
        public TrackerPanel()
        {
            InitializeComponent();
        }

        private bool _compact;

        /// <summary>When true (Workspace), the per-document date boxes are hidden here because they live in the File Dates panel instead.</summary>
        public bool CompactMode
        {
            get => _compact;
            set
            {
                _compact = value;
                var vis = value ? Visibility.Collapsed : Visibility.Visible;
                DatesBlock.Visibility = vis;
                InfoTypeBlock.Visibility = vis;
            }
        }

        private void CalcPlus_Click(object sender, RoutedEventArgs e) { if (Model != null) Model.CalcDayOffset++; }
        private void CalcMinus_Click(object sender, RoutedEventArgs e) { if (Model != null) Model.CalcDayOffset--; }

        private void CalcOffset_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var tb = (TextBox)sender;
            string next = tb.Text.Substring(0, tb.SelectionStart) + e.Text + tb.Text.Substring(tb.SelectionStart + tb.SelectionLength);
            e.Handled = !System.Text.RegularExpressions.Regex.IsMatch(next, @"^-?\d*$");
        }

        private TrackerModel? Model => DataContext as TrackerModel;

        private void RowsGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) != 0 && Keyboard.FocusedElement is not TextBox)
            {
                PasteFromClipboard();
                e.Handled = true;
            }
        }

        // Let the outer ScrollViewer scroll when the grid itself has nothing to scroll.
        private void RowsGrid_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (sender is DataGrid) { /* grid scrolls itself; leave as-is */ }
        }

        private void PasteButton_Click(object sender, RoutedEventArgs e) => PasteFromClipboard();

        private void PasteFromClipboard()
        {
            if (Model == null) return;
            try
            {
                if (!Clipboard.ContainsText())
                {
                    MessageBox.Show("Copy a table from Excel or a web page first, then click Paste Table.", "Paste Table", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                RowsGrid.CommitEdit(DataGridEditingUnit.Row, true);
                int start = RowsGrid.SelectedIndex >= 0 ? RowsGrid.SelectedIndex : 0;
                int n = Model.PasteRows(Clipboard.GetText(), start);
                if (n == 0)
                    MessageBox.Show("No rows could be read. Expected columns: Description, OTR? (Yes/No), Start Date, End Date.", "Paste Table", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not paste: " + ex.Message, "Paste Table", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void AddRowButton_Click(object sender, RoutedEventArgs e) => Model?.Rows.Add(new EmploymentRow());

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            if (Model == null) return;
            if (MessageBox.Show("Clear all dates and employment history?", "Clear Tracker", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                Model.Clear();
        }
    }
}
