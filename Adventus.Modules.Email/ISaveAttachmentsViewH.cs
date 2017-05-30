using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Genesyslab.Desktop.Infrastructure;

namespace Adventus.Modules.Email
{
    public interface ISaveAttachmentsViewH : IView
    {
        ISaveAttachmentsViewModelH Model { get; set; }
    }
}
