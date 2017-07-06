using Genesyslab.Desktop.Infrastructure;

namespace Adventus.Modules.Email
{
    public interface ISaveAttachmentsView : IView
    {
        ISaveAttachmentsViewModel Model { get; set; }
    }
}
