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
using System.IO;
using System.Windows.Media.Media3D;

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
		public bool IsInteractionSelected { get; set; }
		public ContactHistoryView Chv {get;set;}
		public IContactHistoryViewModel Chvm {get;set;}

        public SaveAttachmentsViewH(IObjectContainer container, ISaveAttachmentsViewModelH model)
        {
            this.container = container;
            this.Model = model;
            InitializeComponent();
			IsInteractionSelected = false;
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
			Chv = FindUpVisualTree<ContactHistoryView>(SaveFromHistoryButton);
			//Chv.PropertyChanged += Chv_PropertyChanged;
			((ContactHistoryViewModel)Chv.Model).PropertyChanged += ContactHistoryViewModel_PropertyChanged;
		}

		private void ContactHistoryViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
		{
			this.Dispatcher.Invoke(() =>
			{
				if(e.PropertyName == "InteractionItems")
				{
					if(((ContactHistoryViewModel)Chv.Model).InteractionItems.Count == 0)
					{
						Model.SelectedInteractionId = null;
						Model.SaveButtonVisibilityH = Visibility.Hidden;
						return;
					}
				}
			});
			//else
			//if(e.PropertyName == "SelectedInteractionId")
			//{
			//	Model.SelectedInteractionId = ((ContactHistoryViewModel)Chv.DataContext).SelectedInteractionId;
			//}
			//else
			//{
			//}
		}

		//private void Chv_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
		//{
		//	this.Dispatcher.Invoke(() =>
		//	{
		//		if(e.PropertyName == "SelectedItem")
		//		{
		//			if(Chv.SelectedItem != null)
		//			{
		//				Model.SaveButtonVisibilityH = Visibility.Visible;
		//			}
		//			else
		//			{
		//				Model.SaveButtonVisibilityH = Visibility.Hidden;
		//			}
		//		}
		//	});
		//}

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
					{
					
					//WriteXML(ge);

					//if((string)ge.Action[0].Action == "ActivedThisPanel" && (string)ge.Action[0].Parameters[0] == "MyContactHistory")
					//{
					//	Model.SaveButtonVisibilityH = Visibility.Hidden;
					//	IsInteractionSelected = false;
					//}

					if(ge.Context == Context.ToString())
					{
						if(((ContactHistoryViewModel)Chv.DataContext).SelectedInteractionId != null)
						{
							Model.SelectedInteractionId = ((ContactHistoryViewModel)Chv.DataContext).SelectedInteractionId;
							Model.SaveButtonVisibilityH = Visibility.Visible;
						//if(ge.Target == "ContactHistory")
						//{
							//if((string)ge.Action[0].Action == "LoadInteractionInformation")
							//{

					////if(ge != null && (string)ge.Action[0].Action == "LoadInteractionInformation" && ge.Target == "ContactHistory" && ge.Context == "ContactMain" )
					////if(ge != null && (string)ge.Action[0].Action == "LoadInteractionInformation" && ge.Target == "ContactHistory" && ge.Context == Context.ToString())	// allow saving email from History tab of any parent page, not only contact directory

						//Genesyslab.Desktop.Modules.Contacts.IWInteraction.IWInteractionContent ic =
							//ge.Action[0].Parameters[0] as Genesyslab.Desktop.Modules.Contacts.IWInteraction.IWInteractionContent;
						//Genesyslab.Platform.Contacts.Protocols.ContactServer.InteractionAttributes ia = ic.InteractionAttributes;
						//if(ia.MediaTypeId == "email")
						//{
							//Model.SelectedInteractionId = ia.Id;	// selected interaction id
							//(Model as SaveAttachmentsViewModelBase).Dst = ic.DataSourceType;
							//Model.SaveButtonVisibilityH = Visibility.Visible;
							//IsInteractionSelected = true;
						//}
						//else
						//{
							//Model.SaveButtonVisibilityH = Visibility.Hidden;
							//IsInteractionSelected = false;
						//}

					//}
					//}
					}
					else
					{
						Model.SaveButtonVisibilityH = Visibility.Hidden;
						return;
					}
					}
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

		public void WriteXML(GenericEvent e)  
		{  
	        StreamWriter file = System.IO.File.AppendText("HistoryTabEvents.txt");
			file.WriteLine(String.Format("{0,-20} | {1,-30} | {2,-35} | {3,-30}", DateTime.Now.ToString("HH.mm.ss.ffffff"), "Context: " + e.Context, "Action: " + (string)e.Action[0].Action, "Target: " + e.Target));
			//file.WriteLine("Context: " + e.Context);
			//file.WriteLine("Action: " + (string)e.Action[0].Action);
			//file.WriteLine("Target: " + e.Target);
			file.Close(); 
		}

		public static T FindUpVisualTree<T>(DependencyObject initial) where T : DependencyObject
		{
		    DependencyObject current = initial;
		 
		    while (current != null && current.GetType() != typeof(T))
		    {
				current = LogicalTreeHelper.GetParent(current);
		    }
		    return current as T;   
		}

		//delegate bool ExecuteDelegate(IDictionary<string, object> parameters, IProgressUpdater progressUpdater);
	}
}
