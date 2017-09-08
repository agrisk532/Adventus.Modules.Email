using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using Genesyslab.Desktop.Modules.Windows.Event;
using System.Windows.Threading;
using System.Windows.Media;
using Genesyslab.Desktop.Infrastructure.ViewManager;
using Genesyslab.Desktop.Modules.Contacts.ContactHistory;

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
			IDictionary<string, object> contextDictionary = Context as IDictionary<string, object>;
			Context = (object) contextDictionary["ContactMode"];
			container.Resolve<IViewEventManager>().Subscribe(MyEventHandler2);
			//IViewManager ivm = container.Resolve<IViewManager>();

			//					BundleView bv = this.Model.BundleView;
			//		BundlePartyView bpv = Ivm.GetViewInRegion(bv, "BundlePartyRegion", "BundlePartyView") as BundlePartyView;

			//		object[] v = Ivm.GetAllViewsInRegion(bpv, "PartyRegion");
			//		PartyView pv = v.OfType<PartyView>().First();

			//		object[] v1 = Ivm.GetAllViewsInRegion(pv, "CustomBundlePartyRegion");
			//		InteractionQueueView iqv = v1.OfType<InteractionQueueView>().First();


			//var iv = ivm.GetAllViewsInRegion("ToolbarWorkplaceRegion");
		}

		public void MyEventHandler2(object eventObject)
		{
			// To go to the main thread
			this.Dispatcher.Invoke(() =>
			{
				try
				{
					GenericEvent ge = eventObject as GenericEvent;
					if(ge == null)
						return;
					else
					if(ge.Target != "ContactHistory")
						return;
					else
					//if(ge != null && (string)ge.Action[0].Action == "LoadInteractionInformation" && ge.Target == "ContactHistory" && ge.Context == "ContactMain" )
					if(ge != null && (string)ge.Action[0].Action == "LoadInteractionInformation" && ge.Target == "ContactHistory" && ge.Context == Context.ToString())	// allow saving email from History tab of any parent page, not only contact directory
					{
						Genesyslab.Desktop.Modules.Contacts.IWInteraction.IWInteractionContent ic =
							ge.Action[0].Parameters[0] as Genesyslab.Desktop.Modules.Contacts.IWInteraction.IWInteractionContent;
						Genesyslab.Platform.Contacts.Protocols.ContactServer.InteractionAttributes ia = ic.InteractionAttributes;
						if(ia.MediaTypeId == "email")
						{
							Model.SelectedInteractionId = ia.Id;	// selected interaction id
							(Model as SaveAttachmentsViewModelBase).Dst = ic.DataSourceType;
							Model.SaveButtonVisibilityH = Visibility.Visible;
						}
						else
						{
							Model.SaveButtonVisibilityH = Visibility.Hidden;
						}

					}
					else
					{
						Model.SaveButtonVisibilityH = Visibility.Hidden;
					}
				}
				catch(Exception e)
				{
					MessageBox.Show(string.Format("Exception at processing event {0}", e.Message));
				}
			});
		}

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

		//delegate bool ExecuteDelegate(IDictionary<string, object> parameters, IProgressUpdater progressUpdater);
    }
}
