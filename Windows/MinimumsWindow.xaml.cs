using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using HexxcorpDocumentReviewUtility.Models;
using HexxcorpDocumentReviewUtility.Services;

namespace HexxcorpDocumentReviewUtility.Windows
{
    /// <summary>
    /// Scrollable, read-only modal reference window listing the DOT starter
    /// and user-configured criteria. Content is loaded from
    /// <c>MinimumsData.json</c> via <see cref="MinimumsDataService"/> and can be
    /// edited in-app through the "Edit Minimums..." button without touching code.
    /// </summary>
    public partial class MinimumsWindow : Window
    {
        public MinimumsWindow()
        {
            InitializeComponent();
            LoadSections();
        }

        private void LoadSections()
        {
            MinimumsData data = MinimumsDataService.Load();
            SectionsItemsControl.ItemsSource = data.Sections;
        }

        /// <summary>
        /// Renders a single reference line into a TextBlock's Inlines, supporting
        /// a lightweight <c>**bold**</c> markup and a leading <c>"- "</c> bullet
        /// prefix, so section content can be edited as plain text while still
        /// looking formatted.
        /// </summary>
        private void LineTextBlock_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is not TextBlock textBlock || textBlock.DataContext is not string line)
            {
                return;
            }

            textBlock.Inlines.Clear();

            string text = line;
            if (text.StartsWith("- "))
            {
                textBlock.Inlines.Add(new Run("\u2022 "));
                text = text[2..];
            }

            int cursor = 0;
            while (cursor < text.Length)
            {
                int boldStart = text.IndexOf("**", cursor, System.StringComparison.Ordinal);
                if (boldStart < 0)
                {
                    textBlock.Inlines.Add(new Run(text[cursor..]));
                    break;
                }

                int boldEnd = text.IndexOf("**", boldStart + 2, System.StringComparison.Ordinal);
                if (boldEnd < 0)
                {
                    // Unmatched "**" - treat the rest as plain text rather than silently dropping it.
                    textBlock.Inlines.Add(new Run(text[cursor..]));
                    break;
                }

                if (boldStart > cursor)
                {
                    textBlock.Inlines.Add(new Run(text[cursor..boldStart]));
                }

                textBlock.Inlines.Add(new Run(text[(boldStart + 2)..boldEnd]) { FontWeight = FontWeights.Bold });
                cursor = boldEnd + 2;
            }
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            var editor = new MinimumsEditorWindow { Owner = this };
            if (editor.ShowDialog() == true)
            {
                LoadSections();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
