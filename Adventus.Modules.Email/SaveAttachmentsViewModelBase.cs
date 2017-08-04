using System.Collections.Generic;
using System.ComponentModel;
using Genesyslab.Platform.Contacts.Protocols.ContactServer;

namespace Adventus.Modules.Email
{
	public abstract class SaveAttachmentsViewModelBase : INotifyPropertyChanged
	{
	    
		public List<string> emailPartsPath;   /**< full path on disk of email body and each attachment */
        public bool EmailPartsInfoStored { get; set; }  /**< set this to true after storing message body and all attachment paths */
		public DataSourceType Dst {get; set;} // Main interaction or from archive

		public event PropertyChangedEventHandler PropertyChanged;

        public List<string> EmailPartsPath
        {
            get { return emailPartsPath; }
            set {}
        }

		public SaveAttachmentsViewModelBase()
		{
			emailPartsPath = new List<string>();
			EmailPartsInfoStored = false;
		}

		public void Clear()
		{
			//interaction = null;  // do not clear interaction here. It is used in other parts of code
			EmailPartsPath.Clear();
			EmailPartsInfoStored = false;
		}

        protected void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
	}
}
