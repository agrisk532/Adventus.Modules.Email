using System.Windows;

namespace Adventus.Modules.Email
{
    public interface ISaveAttachmentsViewModelH
    {
        string SelectedInteractionId { get; set; }
		Visibility SaveButtonVisibilityH { get; set; }
    }
}
