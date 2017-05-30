using Genesyslab.Desktop.Infrastructure;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Infrastructure.ViewManager;
using Genesyslab.Desktop.Modules.Windows.Event;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using System.Collections.Generic;
using System;
using System.Windows;

namespace Adventus.Modules.Email
{
/** \class SaveAttachmentsModule
 *  \brief Module for saving email attachments
 */
    public class Module : IModule
    {
        readonly IObjectContainer container;
        readonly IViewManager viewManager;
        readonly ICommandManager commandManager;
		readonly IViewEventManager eventManager;
		private bool isButtonRegisteredInRegion;

/** \brief Initializes a new instance of the SaveAttachmentsModule class.
 *  \param container The container
 *  \param viewManager The view manager
 *  \param commandManager The command manager
 */
        public Module(IObjectContainer container, IViewManager viewManager, ICommandManager commandManager, IViewEventManager eventManager)
        {
            this.container = container;
            this.viewManager = viewManager;
            this.commandManager = commandManager;
			this.eventManager = eventManager;
			this.isButtonRegisteredInRegion = false;
        }

/** \brief Initializes the module
 */
        public void Initialize()
        {
            // Register the view (GUI) "ISaveAttachmentsView" and its behavior counterpart "ISaveAttachmentsViewModel"
            container.RegisterType<ISaveAttachmentsView, SaveAttachmentsView>();
			container.RegisterType<ISendAndSaveAttachmentsView, SendAndSaveAttachmentsView>();
            container.RegisterType<ISaveAttachmentsViewModel, SaveAttachmentsViewModel>();

			// Put the "SaveAttachments" view in the region "BundleCustomButtonRegion" if Condition is true

			viewManager.ViewsByRegionName["BundleCustomButtonRegion"].Add(new ViewActivator()
				{
					ViewType = typeof(ISaveAttachmentsView), ViewName = "SaveAttachments", ActivateView = true, Condition = CheckCondition
				}
			);

			viewManager.ViewsByRegionName["BundleCustomButtonRegion"].Add(new ViewActivator()
				{ 
					ViewType = typeof(ISendAndSaveAttachmentsView), ViewName = "SendAndSaveAttachments", ActivateView = true, Condition = CheckCondition
				}
            );

			// Register commands

			commandManager.CreateChainOfCommandByName("SaveAttachments");
            commandManager.AddCommandToChainOfCommand("SaveAttachments", new List<CommandActivator>()
                {
                    new CommandActivator() { CommandType = typeof(SaveAttachmentsCommand), Name = "SaveAttachments"}
                }
            );

            commandManager.InsertCommandToChainOfCommandBefore("InteractionEmailSend", "Send", new List<CommandActivator>()
                {
                    new CommandActivator() { CommandType = typeof(AttachDataCommand), Name = "AttachData"}
                }
            );

			//viewManager.ViewsByRegionName["ContactHistoryErrorRegion"].Add(new ViewActivator()
			//{
			//	ViewType = typeof(ISaveAttachmentsView),
			//	ViewName = "SaveAttachments",
			//	ActivateView = true,
			//	Condition = CheckCondition2
			//}
			//);

			eventManager.Subscribe(MyEventHandler);
        }

		public void MyEventHandler(object eventObject)
		{
			GenericEvent ge = eventObject as GenericEvent;
			if(!isButtonRegisteredInRegion && ge.Target == "ContactHistory")
			{
				try
				{
					viewManager.ViewsByRegionName["ContactHistoryErrorRegion"].Add(new ViewActivator()
					{
						ViewType = typeof(ISaveAttachmentsView),
						ViewName = "SaveAttachments",
						ActivateView = true,
						Condition = CheckCondition2
					}
					);
					isButtonRegisteredInRegion = true;
				}
				catch(Exception e)
				{
					isButtonRegisteredInRegion = false;
				}
			}
		}

/** \brief View display condition.
 *  executed every time before displaying the view
 *  \param context data from InteractionWorkspace framework
 *  \return true - view will be displayed, false - view will be not displayed
 */
        public bool CheckCondition(ref object context)
        {
            IDictionary<string, object> contextDictionary = context as IDictionary<string, object>;
            if (contextDictionary.ContainsKey("Interaction"))
            {
                IInteraction interaction = contextDictionary["Interaction"] as IInteraction;
                if (interaction != null)
                {
                    if (interaction.EntrepriseInteractionCurrent.IdType.MediaType.ToString() == "Multimedia" &&
                        interaction.EntrepriseInteractionCurrent.IdType.SubMediaType == "email")
                    {
                        return true;  // we store also the email body
                    }
                }
            }    
            return false;
        }

        public bool CheckCondition2(ref object context)
        {
            return true;
        }
    }
}
