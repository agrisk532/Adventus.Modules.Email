using Genesyslab.Desktop.Infrastructure;

namespace Adventus.Modules.Email
{
    public interface ISendAndSaveAttachmentsView : IView
    {
        ISaveAttachmentsViewModel Model { get; set; }
    }
}
