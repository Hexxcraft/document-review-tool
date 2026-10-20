using System.Windows;
using System.Windows.Input;

namespace HexxcorpDocumentReviewUtility.Windows
{
    /// <summary>
    /// Small modal "rename" / single-line text prompt, styled to match the
    /// rest of the app, used by the Workspace document drawer for renaming
    /// files in place.
    /// </summary>
    public partial class TextInputWindow : Window
    {
        public string ResultText { get; private set; } = string.Empty;

        public TextInputWindow(string title, string prompt, string initialValue)
        {
            InitializeComponent();
            Title = title;
            PromptText.Text = prompt;
            InputTextBox.Text = initialValue;
            Loaded += (_, _) =>
            {
                InputTextBox.Focus();
                InputTextBox.SelectAll();
            };
        }

        private void OkButton_Click(object sender, RoutedEventArgs e) => Accept();

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void InputTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Accept();
            }
            else if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }

        private void Accept()
        {
            if (string.IsNullOrWhiteSpace(InputTextBox.Text))
            {
                return;
            }

            ResultText = InputTextBox.Text.Trim();
            DialogResult = true;
            Close();
        }
    }
}
