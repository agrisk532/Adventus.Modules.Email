using System.Windows;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;

namespace Adventus.Modules.Email
{
/** \class SaveAttachmentsViewModel
 *  \brief presentation logic for SaveAttachmentsView
 */
    public class SaveAttachmentsViewModel : SaveAttachmentsViewModelBase, ISaveAttachmentsViewModel
    {
        public IInteraction interaction;        /**< current interaction */
		private Visibility saveButtonVisibility;
		private Visibility sendAndSaveButtonVisibility;

		public Visibility SaveButtonVisibility
		{
	        get
			{
	            return saveButtonVisibility;
			}
			set
			{
			    if (!value.Equals(saveButtonVisibility))
				{
		            saveButtonVisibility = value;
					OnPropertyChanged("SaveButtonVisibility");
				}
			}
		}

		public Visibility SendAndSaveButtonVisibility
		{
	        get
			{
	            return sendAndSaveButtonVisibility;
			}
			set
			{
			    if (!value.Equals(sendAndSaveButtonVisibility))
				{
		            sendAndSaveButtonVisibility = value;
					OnPropertyChanged("SendAndSaveButtonVisibility");
				}
	        }
		}


        public SaveAttachmentsViewModel()
        {
            EmailPartsInfoStored = false;
			SaveButtonVisibility = Visibility.Collapsed;
			SendAndSaveButtonVisibility = Visibility.Collapsed;
        }
		
        public IInteraction Interaction
		{
			get { return interaction; }
			set { if (interaction != value)  interaction = value; }
		}
    }
}
