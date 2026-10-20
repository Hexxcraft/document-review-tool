using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace HexxcorpDocumentReviewUtility.Models
{
    /// <summary>
    /// One labeled section of the DOT reference or user-configured criteria,
    /// containing one or more lines of body text. Each line supports lightweight markup so users
    /// can format entries without editing XAML or code:
    /// <list type="bullet">
    /// <item><description>Wrapping text in <c>**double asterisks**</c> renders it bold.</description></item>
    /// <item><description>A line starting with <c>- </c> renders as a bulleted item.</description></item>
    /// </list>
    /// Implements <see cref="INotifyPropertyChanged"/> solely so the section
    /// list in <c>MinimumsEditorWindow</c> refreshes live as the title is edited.
    /// </summary>
    public class MinimumsSection : INotifyPropertyChanged
    {
        private string _title = string.Empty;

        public string Title
        {
            get => _title;
            set
            {
                if (_title != value)
                {
                    _title = value;
                    OnPropertyChanged();
                }
            }
        }

        public List<string> Lines { get; set; } = new();

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Root object persisted to <c>MinimumsData.json</c> (next to the
    /// application executable). This powers the Review Criteria window and
    /// is editable in-app, so users can add or adjust fields
    /// without needing to touch source code.
    /// </summary>
    public class MinimumsData
    {
        /// <summary>Schema version, allows future format changes without breaking older saves.</summary>
        public int Version { get; set; } = 1;

        public List<MinimumsSection> Sections { get; set; } = new();
    }
}
