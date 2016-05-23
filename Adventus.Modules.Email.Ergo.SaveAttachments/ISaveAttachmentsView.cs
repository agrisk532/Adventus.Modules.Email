using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Genesyslab.Desktop.Infrastructure;

namespace Adventus.Modules.Email.Ergo.SaveAttachments
{
    public interface ISaveAttachmentsView : IView
    {
        ISaveAttachmentsViewModel Model { get; set; }
    }
}
