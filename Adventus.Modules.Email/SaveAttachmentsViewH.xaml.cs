using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using Genesyslab.Desktop.Modules.Windows.Event;
using System.Windows.Threading;

namespace Adventus.Modules.Email
{
	/** \class SaveAttachmentsView
	 *  \brief Interaction logic for SaveAttachmentsView.xaml
	 */
	public partial class SaveAttachmentsViewH : UserControl, ISaveAttachmentsViewH
    {
        readonly IObjectContainer container;
        public object Context { get; set; }
		public ICase Case { get; set; }

        public SaveAttachmentsViewH(IObjectContainer container, ISaveAttachmentsViewModelH model)
        {
            this.container = container;
            this.Model = model;
            InitializeComponent();
        }
        public ISaveAttachmentsViewModelH Model
        {
            get { return this.DataContext as ISaveAttachmentsViewModelH; }
            set { this.DataContext = value; }
        }

/** \brief Executed once, at the view object creation
 */
        public void Create()
        {
			container.Resolve<IViewEventManager>().Subscribe(MyEventHandler2);
		}

		public void MyEventHandler2(object eventObject)
		{
			// To go to the main thread
			this.Dispatcher.Invoke(() =>
			{
				try
				{
					GenericEvent ge = eventObject as GenericEvent;
					if(ge != null && (string)ge.Action[0].Action == "LoadInteractionInformation" && ge.Target == "ContactHistory" && ge.Context == "ContactMain" )
					{
						Genesyslab.Desktop.Modules.Contacts.IWInteraction.IWInteractionContent ic =
						ge.Action[0].Parameters[0] as Genesyslab.Desktop.Modules.Contacts.IWInteraction.IWInteractionContent;
						Genesyslab.Platform.Contacts.Protocols.ContactServer.InteractionAttributes ia = ic.InteractionAttributes;
						if(ia.MediaTypeId == "email")
						{
							Model.SelectedInteractionId = ia.Id;	// selected interaction id
						}
					}
				}
				catch(Exception e)
				{
	
				}
			});
		}
		//public void SAV_InteractionEvent(object sender, EventArgs<IInteraction> e)
		//{
		//      //Add a reference to: Genesyslab.Enterprise.Services.Multimedia.dll 
		//     //and Genesyslab.Enterprise.Model.dll object flag;
		//      IInteraction interaction = e.Value;
		//	  if(interaction.EntrepriseInteractionCurrent.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.Out)
		//	  {
		//			Model.SaveButtonVisibility = Visibility.Collapsed;
		//	  }
		//	  else
		//	  {
		//	  		Model.SaveButtonVisibility = Visibility.Visible;
		//	  }
		//}

/** \brief Executed once, at the view object destruction
 */
        public void Destroy()
        {
			container.Resolve<IViewEventManager>().Unsubscribe(MyEventHandler2);
        }

/** \brief Event handler
 */
        private void SaveAttachmentsButtonH_Click(object sender, RoutedEventArgs e)
        {
            IChainOfCommand Command = container.Resolve<ICommandManager>().GetChainOfCommandByName("SaveAttachments");
            IDictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("Model", Model);
            Command.Execute(parameters);
        }

		delegate bool ExecuteDelegate(IDictionary<string, object> parameters, IProgressUpdater progressUpdater);
    }
}
