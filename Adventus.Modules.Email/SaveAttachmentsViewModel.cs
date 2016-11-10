using System;
using System.Windows;
using System.ComponentModel;
using System.Windows.Input;
using System.Windows.Threading;
using System.Collections.Generic;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;

namespace Adventus.Modules.Email
{
/** \class SaveAttachmentsViewModel
 *  \brief presentation logic for SaveAttachmentsView
 */
    public class SaveAttachmentsViewModel : ISaveAttachmentsViewModel, INotifyPropertyChanged
    {
        public IInteraction interaction;        /**< current interaction */
        public List<string> emailPartsPath = new List<string>();   /**< full path on disk of email body and each attachment */
        public bool EmailPartsInfoStored { get; set; }  /**< set this to true after storing message body and all attachment paths */

		private Visibility saveButtonVisibility;
		private Visibility sendAndSaveButtonVisibility;
		public event PropertyChangedEventHandler PropertyChanged;

		public Visibility SaveButtonVisibility
		{
	        get
			{
	            return saveButtonVisibility;
			}
			set
			{
	            saveButtonVisibility = value;
				OnPropertyChanged("SaveButtonVisibility");
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
	            sendAndSaveButtonVisibility = value;
				OnPropertyChanged("SendAndSaveButtonVisibility");
	        }
		}


        public List<string> EmailPartsPath
        {
            get { return emailPartsPath; }
            set {}
        }
        public SaveAttachmentsViewModel()
        {
            EmailPartsInfoStored = false;
			SaveButtonVisibility = Visibility.Visible;
			SendAndSaveButtonVisibility = Visibility.Hidden;
        }
		
        public IInteraction Interaction
		{
			get { return interaction; }
			set { if (interaction != value)  interaction = value; }
		}

		public void Clear()
		{
			interaction = null;
			EmailPartsPath.Clear();
			EmailPartsInfoStored = false;
		}

		/** \brief send notifications to WPF
 */
        protected void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
