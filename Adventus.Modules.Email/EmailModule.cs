using Genesyslab.Desktop.Infrastructure;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Infrastructure.ViewManager;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using System.Collections.Generic;

namespace Adventus.Modules.Email
{
/** \class SaveAttachmentsModule
 *  \brief Module for saving email attachments
 */
    public class AdventusEmailModule : IModule
    {
        readonly IObjectContainer container;
        readonly IViewManager viewManager;
        readonly ICommandManager commandManager;

/** \brief Initializes a new instance of the SaveAttachmentsModule class.
 *  \param container The container
 *  \param viewManager The view manager
 *  \param commandManager The command manager
 */
        public AdventusEmailModule(IObjectContainer container, IViewManager viewManager, ICommandManager commandManager)
        {
            this.container = container;
            this.viewManager = viewManager;
            this.commandManager = commandManager;
        }

/** \brief Initializes the module
 */
        public void Initialize()
        {
            // Register the view (GUI) "ISaveAttachmentsView" and its behavior counterpart "ISaveAttachmentsViewModel"
            container.RegisterType<ISaveAttachmentsView, SaveAttachmentsView>();
            container.RegisterType<ISaveAttachmentsViewModel, SaveAttachmentsViewModel>();

            // Put the "SaveAttachments" view in the region "BundleCustomButtonRegion" if Condition is true
            viewManager.ViewsByRegionName["BundleCustomButtonRegion"].Add(new ViewActivator() { ViewType = typeof(ISaveAttachmentsView), ViewName = "SaveAttachments", ActivateView = true,
                Condition = CheckCondition
/* delegate code moved to the CheckCondition() method to make Visual Studio debugging easier
                Condition = delegate(ref object context)
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
                                return true;
                            }
                        }
                    }    
                    return false;
                }
 */
            }
            );
            // register commands
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
//                        IInteractionEmail interactionEmail = interaction as IInteractionEmail;
//                        if ((interactionEmail.EntrepriseEmailAttachments != null) && (interactionEmail.EntrepriseEmailAttachments.Count > 0))
//                        {
//                            return true;
//                        }
                    }
                }
            }    
            return false;
        }
    }
}
