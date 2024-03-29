﻿using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Enterprise.Model.ServiceModel;
using Genesyslab.Enterprise.Services;
using Genesyslab.Enterprise.Model.Contact;
using Genesyslab.Platform.Commons.Protocols;
using Genesyslab.Desktop.Modules.Windows.Interactions;

namespace Adventus.Modules.Email
{
/** \class SaveAttachmentsView
 *  \brief Interaction logic for SaveAttachmentsView.xaml
 */
    public partial class SendAndSaveAttachmentsView : UserControl, ISendAndSaveAttachmentsView
    {
        readonly IObjectContainer container;
        public object Context { get; set; }
		public ICase Case { get; set; }

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
			Case = contextDictionary["Case"] as ICase;
            //Model.Interaction = contextDictionary.TryGetValue("Interaction") as IInteraction;
            //IInteractionEmail interactionEmail = Model.Interaction as IInteractionEmail;
            //if (interactionEmail == null)
            //{
            //    MessageBox.Show("Interaction is not of IInteractionEmail type");
            //}

			//container.Resolve<IInteractionManager>().InteractionEvent += 
			//	new System.EventHandler<EventArgs<IInteraction>> (SASAV_InteractionEvent);
			container.Resolve<IInteractionsWindowController>().InteractionViewCreated += SendAndSaveAttachmentsView_InteractionViewCreated;
        }

		private void SendAndSaveAttachmentsView_InteractionViewCreated(object sender, InteractionViewEventArgs e)
		{
			//IInteractionEmail eventInteractionEmail = e.Interaction as IInteractionEmail;
			//if(eventInteractionEmail == null)
			//{
			//	return;		// ignore non-email type interactions in changing custom email buttons
			//}
			//else
			if(e.Interaction.CaseId == Case.CaseId)
			{
			//Model.Interaction = eventInteractionEmail;
			//IInteractionEmail modelInteractionEmail = Model.Interaction as IInteractionEmail;
			//Model.Interaction = eventInteractionEmail;
			Model.Interaction = e.Interaction;

			//if(eventInteractionEmail.EntrepriseEmailInteractionCurrent.Id		== modelInteractionEmail.EntrepriseEmailInteractionCurrent.Id ||
			//   eventInteractionEmail.EntrepriseEmailInteractionCurrent.ParentID	== modelInteractionEmail.EntrepriseEmailInteractionCurrent.Id)
			//{
				if(e.Interaction.EntrepriseInteractionCurrent.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.Out)
				{
					//Model.SaveButtonVisibility = Visibility.Collapsed;
					Model.SendAndSaveButtonVisibility = Visibility.Visible;
				}
				else
				if(e.Interaction.EntrepriseInteractionCurrent.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.In)
				{
					//Model.SaveButtonVisibility = Visibility.Visible;
					Model.SendAndSaveButtonVisibility = Visibility.Collapsed;
				}
			}
		}



		//public void SASAV_InteractionEvent(object sender, EventArgs<IInteraction> e)
		//{
		//      //Add a reference to: Genesyslab.Enterprise.Services.Multimedia.dll 
		//     //and Genesyslab.Enterprise.Model.dll object flag;
		//      IInteraction interaction = e.Value;
		//	  if(interaction.EntrepriseInteractionCurrent.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.In)
		//	  {
		//			Model.SendAndSaveButtonVisibility = Visibility.Collapsed;
		//	  }
		//	  else
		//	  {
		//	  		Model.SendAndSaveButtonVisibility = Visibility.Visible;
		//	  }
		//}

/** \brief Executed once, at the view object destruction
 */
        public void Destroy()
        {
			//container.Resolve<IInteractionManager>().InteractionEvent -= 
			//	new System.EventHandler<EventArgs<IInteraction>> (SASAV_InteractionEvent);
			container.Resolve<IInteractionsWindowController>().InteractionViewCreated -= SendAndSaveAttachmentsView_InteractionViewCreated;
        }

/** \brief Event handler
 */
        private void SendAndSaveAttachmentsButton_Click(object sender, RoutedEventArgs e)
        {
	        //IDictionary<string, object> contextDictionary = Context as IDictionary<string, object>;
         //   IInteraction interaction_ = contextDictionary.TryGetValue("Interaction") as IInteraction;
         //   IInteractionEmail interactionEmail_ = interaction_ as IInteractionEmail;
			IInteraction interaction = Model.Interaction;
			IInteractionEmail interactionEmail = interaction as IInteractionEmail;

			//if(interaction.Equals(interaction_))
			//{
			//	MessageBox.Show("Interactions are equal");
			//}
			//else
			//{
			//	MessageBox.Show("Interactions are not equal");
			//}

			IChainOfCommand Command;
            IDictionary<string, object> parameters = new Dictionary<string, object>();
			ICommandManager commandManager = container.Resolve<ICommandManager>();

			if (interactionEmail.EntrepriseEmailInteractionCurrent.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.Out && 
				interactionEmail.EntrepriseEmailInteractionCurrent.IdType.Subtype != "OutboundNew" &&
				interactionEmail.EntrepriseEmailInteractionCurrent.IdType.Subtype != "OutboundReply")
			{

				// add attachments from the parent interaction
				string InteractionParentID = interactionEmail.EntrepriseEmailInteractionCurrent.ParentID;

				if (String.IsNullOrEmpty(InteractionParentID))
				{
					MessageBox.Show("Interaction ParentID is null. Cannot add parent interaction attachments", "Attention");
					return;
				}
				else
				{
					// add attachments from ParentID interaction to this interaction
					Genesyslab.Enterprise.Services.IContactService service = container.Resolve<IEnterpriseServiceProvider>().Resolve<IContactService>("contactService");
					Genesyslab.Desktop.Modules.Core.SDK.Contact.IContactService service2 = container.Resolve<Genesyslab.Desktop.Modules.Core.SDK.Contact.IContactService>();
					Genesyslab.Enterprise.Model.Channel.IClientChannel channel = container.Resolve<Genesyslab.Desktop.Modules.Core.SDK.Protocol.IChannelManager>().Register(service2.UCSApp, "IW@ContactService");

					ICollection<IAttachment> attachments = new List<IAttachment>();
					//ICollection<IAttachment> attachments2 = new List<IAttachment>();
					if ((channel != null) && (channel.State == ChannelState.Opened))
					{
						attachments = service.GetAttachments(channel, InteractionParentID, false);  // without attachment body
					}
					if (attachments.Count > 0)
					{
						foreach (IAttachment attachment in attachments)
						{
							if (attachment != null)
							{
								service.AddAttachment(channel, interaction.EntrepriseInteractionCurrent.Id, attachment.Id);
							}
						}
					}
					//attachments2 = service.GetAttachments(channel, interaction.EntrepriseInteractionCurrent.Id, true);  // with attachment body
				}
				Command = container.Resolve<ICommandManager>().GetChainOfCommandByName("InteractionEmailSave");
				parameters.Clear();
				parameters.Add("CommandParameter", interactionEmail);
				Command.Execute(parameters);
			}

			Command = container.Resolve<ICommandManager>().GetChainOfCommandByName("InteractionEmailSend");
			parameters.Clear();
            parameters.Add("CommandParameter", interaction);
            Command.Execute(parameters);
		
			// save email to filesystem. Binary contents of the outgoing email at this point is not available from API. Email created by an agent has not yet traveled the Business Process.
			// available are only email parts created by an agent. .eml file has to be assembled from the email parts.
			Command = container.Resolve<ICommandManager>().GetChainOfCommandByName("SaveAttachments");
			parameters.Clear();
            parameters.Add("Model", Model);
            Command.Execute(parameters);
        }
    }
}
