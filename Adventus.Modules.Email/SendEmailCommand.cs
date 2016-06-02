using System.Collections.Generic;
using System.Windows;
using System.Windows.Threading;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email;
using Genesyslab.Platform.Commons.Logging;

namespace Adventus.Modules.Email
{
/** \class SendEmailCommandBase
 *  \brief base class for commands
 */
    public class SendEmailCommandBase : IElementOfCommand
    {
        public readonly IObjectContainer container;
        public readonly ILogger log;
        public string Name { get; set; }
        public IInteraction interaction { get; set; }
        public IInteractionEmail interactionEmail { get; set; }

        public SendEmailCommandBase(IObjectContainer container, ILogger logger)
        {
            this.container = container;
            this.log = logger;
       }

        public virtual bool Execute(IDictionary<string, object> parameters, IProgressUpdater progress)
        {
            return false;
        }


/** \brief Prints message in the logfile
 *  \param msg a message
 */
        public void ErrorMessage(string msg)
        {
            log.Info(msg);
        }

        protected delegate bool ExecuteDelegate(IDictionary<string, object> parameters, IProgressUpdater progressUpdater);
    }

/** \class AttachDataCommand
 *  \brief Attaches data to email interaction before sending email
 */
    public class AttachDataCommand : SendEmailCommandBase
    {
        /** \brief Constructor, uses dependency injection
         */
        public AttachDataCommand(IObjectContainer container, ILogger logger) : base(container, logger)
        {
        }

        /** \brief Command implementation
         *  Outbound email interaction must be saved before calling this command.
         *  Otherwise the reply email text typed by agent and added attachments will be not available.
         *  \param parameters - input data
         *  \param progress
         *  \return true to stop execution of the command chain; otherwise, false.
         */
        public override bool Execute(IDictionary<string, object> parameters, IProgressUpdater progress)
        {
            log.Info("*** AttachDataCommand() entered ***");
            // To go to the main thread
            if (Application.Current.Dispatcher != null && !Application.Current.Dispatcher.CheckAccess())
            {
                object result = Application.Current.Dispatcher.Invoke(DispatcherPriority.Send, new ExecuteDelegate(Execute), parameters, progress);
                return (bool)result;
            }
            else
            {
                this.interaction = parameters["CommandParameter"] as IInteraction;
                interactionEmail = interaction as IInteractionEmail;

                if (interaction == null)
                {
                    ErrorMessage("Interaction is NULL");
                    return true; // stop execution of the command chain
                }
                else
                if (interactionEmail == null)
                {
                    ErrorMessage("Interaction is not of IInteractionEmail type");
                    return true;
                }
                else
                {
                    interaction.SetAttachedData("_Sending", 1); 
					return false;
				}
            }
        }
    }
}

