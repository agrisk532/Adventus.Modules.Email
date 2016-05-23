using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;

namespace Adventus.Modules.Email.Ergo.SaveAttachments
{
    public interface ISaveAttachmentsViewModel
    {
        IInteraction Interaction { get; set; }
    }
}
