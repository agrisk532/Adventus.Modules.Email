using System.Windows;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;

namespace Adventus.Modules.Email
{
    public interface ISaveAttachmentsViewModel
    {
        IInteraction Interaction { get; set; }
		Visibility SaveButtonVisibility { get; set; }
		Visibility SendAndSaveButtonVisibility { get; set; }
    }
}
