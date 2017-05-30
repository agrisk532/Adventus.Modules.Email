﻿using System.Windows;
using System.ComponentModel;
using System.Collections.Generic;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;

namespace Adventus.Modules.Email
{
/** \class SaveAttachmentsViewModel
 *  \brief presentation logic for SaveAttachmentsView
 */
    public class SaveAttachmentsViewModelH : ISaveAttachmentsViewModelH, INotifyPropertyChanged
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


        public List<string> EmailPartsPath
        {
            get { return emailPartsPath; }
            set {}
        }
        public SaveAttachmentsViewModelH()
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

		public void Clear()
		{
			//interaction = null;  // do not clear interaction here. It is used in other parts of code
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