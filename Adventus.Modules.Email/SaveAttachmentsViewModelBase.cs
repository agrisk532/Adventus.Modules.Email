using System.Collections.Generic;
using System.ComponentModel;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;

namespace Adventus.Modules.Email
{
	public class SaveAttachmentsViewModelBase : INotifyPropertyChanged
	{
	    public List<string> emailPartsPath = new List<string>();   /**< full path on disk of email body and each attachment */
        public bool EmailPartsInfoStored { get; set; }  /**< set this to true after storing message body and all attachment paths */

		public event PropertyChangedEventHandler PropertyChanged;

        public List<string> EmailPartsPath
        {
            get { return emailPartsPath; }
            set {}
        }

		public IInteraction Interaction { get; internal set; }

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
