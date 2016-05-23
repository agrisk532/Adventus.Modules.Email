using System;
using System.ComponentModel;
using System.Windows.Input;
using System.Windows.Threading;
using System.Collections.Generic;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;

namespace Adventus.Modules.Email.Ergo.SaveAttachments
{
/** \class SaveAttachmentsViewModel
 *  \brief presentation logic for SaveAttachmentsView
 */
    public class SaveAttachmentsViewModel : ISaveAttachmentsViewModel
    {
        IInteraction interaction;        /**< current interaction */
        List<string> emailPartsPath = new List<string>();   /**< full path on disk of email body and each attachment */
        public bool EmailPartsInfoStored { get; set; }  /**< set this to true after storing message body and all attachment paths */

        public List<string> EmailPartsPath
        {
            get { return emailPartsPath; }
            set {}
        }
        public SaveAttachmentsViewModel()
        {
            EmailPartsInfoStored = false;
        }
		
        public IInteraction Interaction
		{
			get { return interaction; }
			set { if (interaction != value)  interaction = value; }
		}
    }
}
