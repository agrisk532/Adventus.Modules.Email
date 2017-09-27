using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.Windows.Event;
using System.Windows.Media;
using Genesyslab.Desktop.Modules.Contacts.ContactHistory;
using System.IO;
using Genesyslab.Desktop.WPFCommon.Controls;
using Genesyslab.Desktop.Modules.Windows.Interactions;
using Genesyslab.Desktop.Modules.Contacts.ContactDetail;

namespace Adventus.Modules.Email
{
	/** \class SaveAttachmentsView
	 *  \brief Interaction logic for SaveAttachmentsView.xaml
	 */
	public partial class SaveAttachmentsViewH : UserControl, ISaveAttachmentsViewH
    {
        readonly IObjectContainer container;
        public object Context { get; set; }
		public ContactHistoryView Chv {get; set;}
		//public ContactHistoryViewModel Chvm {get;set;}
		//IInteractionItemViewModel SelectedItem {get;set;}
		// Config server parameter for attached (user) data parameter name. It is used to enable/disable confidential email display
		private const string CONFIG_OPTION_CONFIDENTIAL_EMAIL_ATTACHED_DATA_SECTION_NAME = "custom-email-content-save";
		private const string CONFIG_OPTION_CONFIDENTIAL_EMAIL_ATTACHED_DATA_PARAMETER_NAME = "email-content-display-confidential-attached-data-parameter-name";
		private const string CONFIG_OPTION_CONFIDENTIAL_EMAIL_ATTACHED_DATA_PARAMETER_VALUE = "email-content-display-confidential-attached-data-parameter-value";
		private const string METHOD_NAME = "Adventus.Modules.Email.SaveAttachmentsViewH(): ";	

		private string ConfidentialInfoParamName  {get; set;}
		private string ConfidentialInfoParamValue {get; set;}

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
			container.Resolve<IViewEventManager>().Subscribe(IViewEventManager_EventHandler);
			Chv = FindUpVisualTree<ContactHistoryView>(SaveFromHistoryButton);
			ConfidentialInfoParamName = Util.GetConfigurationOption(CONFIG_OPTION_CONFIDENTIAL_EMAIL_ATTACHED_DATA_SECTION_NAME,
					CONFIG_OPTION_CONFIDENTIAL_EMAIL_ATTACHED_DATA_PARAMETER_NAME, container, METHOD_NAME);
			ConfidentialInfoParamValue = Util.GetConfigurationOption(CONFIG_OPTION_CONFIDENTIAL_EMAIL_ATTACHED_DATA_SECTION_NAME,
					CONFIG_OPTION_CONFIDENTIAL_EMAIL_ATTACHED_DATA_PARAMETER_VALUE, container, METHOD_NAME);

			//Chvm = (ContactHistoryViewModel)Chv.DataContext;
			//Chv.PropertyChanged += Chv_PropertyChanged;
			//Chvm.PropertyChanged += ContactHistoryViewModel_PropertyChanged;
		}

		//private void ContactHistoryViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
		//{
		//	this.Dispatcher.Invoke(() =>
		//	{
		//		if(e.PropertyName == "InteractionItems")
		//		{
		//			// for some reason this condition is always true
		//			//if(((ContactHistoryViewModel)Chv.Model).InteractionItems.Count == 0)
		//			//{
		//			//	Model.SelectedInteractionId = null;
		//			//	Model.SaveButtonVisibilityH = Visibility.Hidden;
		//			//	return;
		//			//}
		//			Model.SelectedInteractionId = null;
		//			Model.SaveButtonVisibilityH = Visibility.Hidden;
		//		}
		//		else
		//		if(e.PropertyName == "SelectedInteractionId")
		//		{
		//			Model.SelectedInteractionId = Chvm.SelectedInteractionId;
		//			//var i = GetChildOfType<ListView>(Chv); // this line works if needed
		//			string type = Chv.GetTypeForInteractionId(Chvm.SelectedInteractionId);
		//			Model.SaveButtonVisibilityH = (type == "email") ? Visibility.Visible : Visibility.Hidden;
		//		}
		//		else
		//		{
		//			Model.SaveButtonVisibilityH = Visibility.Hidden;
		//		}
		//	});
		//	//else
		//	//if(e.PropertyName == "SelectedInteractionId")
		//	//{
		//	//	Model.SelectedInteractionId = ((ContactHistoryViewModel)Chv.DataContext).SelectedInteractionId;
		//	//}
		//	//else
		//	//{
		//	//}
		//}

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

		public void IViewEventManager_EventHandler(object eventObject)
		{
			// To go to the main thread
			this.Dispatcher.Invoke(() =>
			{
				try
				{
					GenericEvent ge = eventObject as GenericEvent;
					if(ge == null)
						return;
					else	// ge != null
					{
					
					//WriteXML(ge);

						if(ge.Context == Context.ToString())	// event for corresponding instance of SaveAttachmentsViewH. In total there are 3 instances.
						{
							string s = ((ContactHistoryViewModel)Chv.DataContext).SelectedInteractionId;
							if(s != null)
							{
								Model.SelectedInteractionId = s;
								if (ge.Target == "ContactHistory")
								{
									if ((string)ge.Action[0].Action == "LoadInteractionInformation")
									{
	
										//if(ge != null && (string)ge.Action[0].Action == "LoadInteractionInformation" && ge.Target == "ContactHistory" && ge.Context == "ContactMain" )
										//if(ge != null && (string)ge.Action[0].Action == "LoadInteractionInformation" && ge.Target == "ContactHistory" && ge.Context == Context.ToString())	// allow saving email from History tab of any parent page, not only contact directory
	
										Genesyslab.Desktop.Modules.Contacts.IWInteraction.IWInteractionContent ic =
											ge.Action[0].Parameters[0] as Genesyslab.Desktop.Modules.Contacts.IWInteraction.IWInteractionContent;
										Genesyslab.Platform.Contacts.Protocols.ContactServer.InteractionAttributes ia = ic.InteractionAttributes;
										if (ia.MediaTypeId == "email")
										{
											//SortableTabControl Stc = GetChildOfType<SortableTabControl>(Chv);
											SortableTabControl Stc = getTabControl();
											if(Stc != null)
											{
												DockPanel dp = getDockPanelInteractionActions();
												Model.SelectedInteractionId = ia.Id;    // selected interaction id
												(Model as SaveAttachmentsViewModelBase).Dst = ic.DataSourceType;
												string attachedData = (string)ia.AllAttributes[ConfidentialInfoParamName] ?? String.Empty;
												if(attachedData != String.Empty && attachedData == ConfidentialInfoParamValue)
												{
													foreach (UserControl uc in Stc.Items)	// hide info in tab control
													{
														//if (uc is IStaticCaseDataView) { uc.Visibility = Visibility.Hidden;}
														if (uc is INotepadView || uc is IContactDetailView)
														{
															uc.Visibility = Visibility.Hidden;
														}
													}
													dp.Visibility = Visibility.Hidden;	// hide dockPanelInteractionActions
													Model.SaveButtonVisibilityH = Visibility.Hidden;	// hide SaveAttachmentsViewH button
												}
												else
												{
											        foreach(UserControl uc in Stc.Items)
													{
														uc.Visibility = Visibility.Visible;
													}
													dp.Visibility = Visibility.Visible;
													Model.SaveButtonVisibilityH = Visibility.Visible;
												}
											}
											else	// Stc == null
											{
												Model.SaveButtonVisibilityH = Visibility.Hidden;
											}
										}
										else	// ia.MediaTypeId != "email"
										{
											// ia.MediaTypeId != "email"
											Model.SaveButtonVisibilityH = Visibility.Hidden;
										}
									}
									else	// (string)ge.Action[0].Action != "LoadInteractionInformation"
									{
									}
								}
								else	// ge.Target != "ContactHistory"
								{
									Model.SaveButtonVisibilityH = Visibility.Hidden;
									return;
								}
							}
							else	// (ContactHistoryViewModel)Chv.DataContext).SelectedInteractionId == null
							{
								// There is no selected interaction
							}
						}
						else	// ge.Context != Context.ToString(). This is event for another instance of SaveAttachmentsViewH. There 3 instances in total
						{
						}
					}
				}
				catch(Exception e)
				{
					MessageBox.Show(string.Format("Exception at processing event {0}", e.Message));
				}
			});
		}

		private DockPanel getDockPanelInteractionActions()
		{
			return FindChild(Chv, child =>
			{
				var dockPanel = child as DockPanel;
				if (dockPanel != null && dockPanel.Name == "dockPanelInteractionActions")
					return true;
				else
					return false;
			}) as DockPanel;
		}

		private SortableTabControl getTabControl()
		{
			return FindChild(Chv, child =>
			{
				var stc = child as SortableTabControl;
				if (stc != null && stc.Name == "tabControlContactHistoryMultiViews")
					return true;
				else
					return false;
			}) as SortableTabControl;
		}

		/** \brief Executed once, at the view object destruction
		 */
		public void Destroy()
        {
			container.Resolve<IViewEventManager>().Unsubscribe(IViewEventManager_EventHandler);
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
			file.Close(); 
		}

		// Visual tree methods

		public static T FindUpVisualTree<T>(DependencyObject initial) where T : DependencyObject
		{
		    DependencyObject current = initial;
		 
		    while (current != null && current.GetType() != typeof(T))
		    {
				current = LogicalTreeHelper.GetParent(current);
		    }
		    return current as T;   
		}

		public static T GetChildOfType<T>(DependencyObject depObj) where T : DependencyObject
		{
		    if (depObj == null) return null;
		
		    for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
		    {
		        var child = VisualTreeHelper.GetChild(depObj, i);
		
		        var result = (child as T) ?? GetChildOfType<T>(child);
		        if (result != null) return result;
		    }
		    return null;
		}

	
		public static DependencyObject FindChild(DependencyObject parent, Func<DependencyObject, bool> predicate)
		{
		    if (parent == null) return null;
		
		    int childrenCount = VisualTreeHelper.GetChildrenCount(parent);
		    for (int i = 0; i < childrenCount; i++)
		    {
		        var child = VisualTreeHelper.GetChild(parent, i);
		
		        if (predicate(child))
		        {
		            return child;
		        }
		        else
		        {
		            var foundChild = FindChild(child, predicate);
		            if (foundChild != null)
		                return foundChild;
		        }
		    }
		
		    return null;
		}
	}
}
