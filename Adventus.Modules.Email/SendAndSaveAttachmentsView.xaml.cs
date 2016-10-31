using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Genesyslab.Desktop.Infrastructure;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email;
using Genesyslab.Desktop.Modules.OpenMedia.Windows.Interactions.MediaView.Email.InteractionInboundEmailView;
using Genesyslab.Desktop.WPFCommon;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;

namespace Adventus.Modules.Email
{
/** \class SaveAttachmentsView
 *  \brief Interaction logic for SaveAttachmentsView.xaml
 */
    public partial class SendAndSaveAttachmentsView : UserControl, ISendAndSaveAttachmentsView
    {
        readonly IObjectContainer container;
        public object Context { get; set; }

        public SendAndSaveAttachmentsView(IObjectContainer container, ISaveAttachmentsViewModel model)
        {
            this.container = container;
            this.Model = model;
            InitializeComponent();
        }
        public ISaveAttachmentsViewModel Model
        {
            get { return this.DataContext as ISaveAttachmentsViewModel; }
            set { this.DataContext = value; }
        }

/** \brief Executed once, at the view object creation
 */
        public void Create()
        {
            IDictionary<string, object> contextDictionary = Context as IDictionary<string, object>;
            Model.Interaction = contextDictionary.TryGetValue("Interaction") as IInteraction;
            IInteractionEmail interactionEmail = Model.Interaction as IInteractionEmail;
            if (interactionEmail == null)
            {
                MessageBox.Show("Interaction is not of IInteractionEmail type");
            }
        }

/** \brief Executed once, at the view object destruction
 */
        public void Destroy()
        {
        }

/** \brief Event handler
 */
        private void SendAndSaveAttachmentsButton_Click(object sender, RoutedEventArgs e)
        {
		// send email
            IDictionary<string, object> parameters = new Dictionary<string, object>();
			IChainOfCommand Command = container.Resolve<ICommandManager>().GetChainOfCommandByName("InteractionEmailSend");
			parameters.Clear();
            parameters.Add("CommandParameter", Model.Interaction);
            Command.Execute(parameters);
		// save email to filesystem
            Command = container.Resolve<ICommandManager>().GetChainOfCommandByName("SaveAttachments");
			parameters.Clear();
            parameters.Add("Model", Model);
            Command.Execute(parameters);
        }
    }
}
