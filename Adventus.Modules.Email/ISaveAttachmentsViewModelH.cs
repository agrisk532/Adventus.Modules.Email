using System;
using System.Windows;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;

namespace Adventus.Modules.Email
{
    public interface ISaveAttachmentsViewModelH
    {
        IInteraction Interaction { get; set; }
		void Clear();
		Visibility SaveButtonVisibility { get; set; }
		Visibility SendAndSaveButtonVisibility { get; set; }
    }
}
