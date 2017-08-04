using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.Windows.Interactions;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using Genesyslab.Desktop.Infrastructure;

namespace Adventus.Modules.Email
{
	/** \class SaveAttachmentsView
	 *  \brief Interaction logic for SaveAttachmentsView.xaml
	 */
	public partial class SaveAttachmentsView : UserControl, ISaveAttachmentsView
    {
        readonly IObjectContainer container;
        public object Context { get; set; }
		public ICase Case { get; set; }

        public SaveAttachmentsView(IObjectContainer container, ISaveAttachmentsViewModel model)
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
				//		 new System.EventHandler<EventArgs<IInteraction>>(SAV_InteractionEvent);
			container.Resolve<IInteractionsWindowController>().InteractionViewCreated += SaveAttachmentsView_InteractionViewCreated;
		}

		private void SaveAttachmentsView_InteractionViewCreated(object sender, InteractionViewEventArgs e)
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
			
				Model.Interaction = e.Interaction;
				(Model as SaveAttachmentsViewModelBase).Dst = Genesyslab.Platform.Contacts.Protocols.ContactServer.DataSourceType.Main;

			//if(eventInteractionEmail.EntrepriseEmailInteractionCurrent.Id		== modelInteractionEmail.EntrepriseEmailInteractionCurrent.Id ||
			//   eventInteractionEmail.EntrepriseEmailInteractionCurrent.ParentID	== modelInteractionEmail.EntrepriseEmailInteractionCurrent.Id)
			//{
				if(e.Interaction.EntrepriseInteractionCurrent.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.Out)
				{
					Model.SaveButtonVisibility = Visibility.Collapsed;
					//Model.SendAndSaveButtonVisibility = Visibility.Visible;
				}
				else
				if(e.Interaction.EntrepriseInteractionCurrent.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.In)
				{
					Model.SaveButtonVisibility = Visibility.Visible;
					//Model.SendAndSaveButtonVisibility = Visibility.Collapsed;
				}
			}
		}

		//public void SAV_InteractionEvent(object sender, EventArgs<IInteraction> e)
		//{
		//	//Add a reference to: Genesyslab.Enterprise.Services.Multimedia.dll 
		//	//and Genesyslab.Enterprise.Model.dll object flag;
		//	IInteraction interaction = e.Value;
		//	if (interaction.EntrepriseInteractionCurrent.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.Out)
		//	{
		//		Model.SaveButtonVisibility = Visibility.Collapsed;
		//	}
		//	else
		//	{
		//		Model.SaveButtonVisibility = Visibility.Visible;
		//	}
		//}

		/** \brief Executed once, at the view object destruction
		 */
		public void Destroy()
        {
			//container.Resolve<IInteractionManager>().InteractionEvent -= 
				//new System.EventHandler<EventArgs<IInteraction>> (SAV_InteractionEvent);
			container.Resolve<IInteractionsWindowController>().InteractionViewCreated -= SaveAttachmentsView_InteractionViewCreated;
        }

/** \brief Event handler
 */
        private void SaveAttachmentsButton_Click(object sender, RoutedEventArgs e)
        {
            IChainOfCommand Command = container.Resolve<ICommandManager>().GetChainOfCommandByName("SaveAttachments");
            IDictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("Model", Model);
            Command.Execute(parameters);
        }
    }
}
