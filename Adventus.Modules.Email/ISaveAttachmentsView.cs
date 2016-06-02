using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Genesyslab.Desktop.Infrastructure;

namespace Adventus.Modules.Email
{
    public interface ISaveAttachmentsView : IView
    {
        ISaveAttachmentsViewModel Model { get; set; }
    }
}
