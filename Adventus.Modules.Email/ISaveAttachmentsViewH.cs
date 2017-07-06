using Genesyslab.Desktop.Infrastructure;

namespace Adventus.Modules.Email
{
    public interface ISaveAttachmentsViewH : IView
    {
        ISaveAttachmentsViewModelH Model { get; set; }
    }
}
