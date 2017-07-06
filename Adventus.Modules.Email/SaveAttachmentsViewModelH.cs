using System.Windows;

namespace Adventus.Modules.Email
{
/** \class SaveAttachmentsViewModel
 *  \brief presentation logic for SaveAttachmentsView
 */
    public class SaveAttachmentsViewModelH : SaveAttachmentsViewModelBase, ISaveAttachmentsViewModelH
    {
		public string SelectedInteractionId {get; set;} // selected in history interaction id

		private Visibility saveButtonVisibilityH;
		

		public Visibility SaveButtonVisibilityH
		{
	        get
			{
	            return saveButtonVisibilityH;
			}
			set
			{
			    if (!value.Equals(saveButtonVisibilityH))
				{
		            saveButtonVisibilityH = value;
					OnPropertyChanged("SaveButtonVisibilityH");
				}
			}
		}

        public SaveAttachmentsViewModelH()
        {
			//SaveButtonVisibilityH = Visibility.Collapsed;
			SaveButtonVisibilityH = Visibility.Visible;
        }
		
		/** \brief send notifications to WPF
 */
    }
}
