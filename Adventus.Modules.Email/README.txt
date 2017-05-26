return this.Interaction.ViewData["OutboundEmailView_Attachments"] as ObservableCollection<IAttachmentGraphic>;

interaction.ViewData["OutboundEmailView_Attachments"]
UCS port: 6120
ucsConnection = new UniversalContactServerProtocol(new Endpoint("eS_UniversalContactServer", "genesys1", 6120));

eml format works with .net ver 4.5. That is due to MailMessage.Send() function.



1)Email sending sequence:
InteractionToolBarButton buttonSend
buttonSend_click
ExecutedSendCommand((object) this.buttonSend, (ExecutedRoutedEventArgs) null);

Utils.ExecuteAsynchronousCommand(this.Model.SendCommand, (IDictionary<string, object>) new Dictionary<string, object>()
      {
        {
          "CommandParameter",
          (object) this.Model.Interaction
        }
      }, sender as UIElement);
    }

this.SendCommand = commandManager.GetChainOfCommandByName("InteractionEmailSend");
CommandManager.AddCommandToChainOfCommand("InteractionEmailSend", (IList<CommandActivator>) new List<CommandActivator>() {
		...
        new CommandActivator()
        {
          CommandType = typeof (SetAttachedDadaInformationSaveEmailCommand),
          Name = "Save"
        },
        new CommandActivator()
        {
          CommandType = typeof (CheckEmailFieldsBeforeSendEmailCommand),
          Name = "CheckEmailFieldsBeforeSend"
        },
        new CommandActivator()
        {
          CommandType = typeof (SendEmailCommand),
          Name = "Send"
        }
}

SendEmailCommand uses Genesyslab.Enterprise.Services.EmailService:
Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email.SendEmailCommand.cs -> this.eMailService.Update(interactionOutboundEmail.EntrepriseEmailInteractionCurrent);
Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email.SendEmailCommand.cs -> this.eMailService.Send(interactionOutboundEmail.EntrepriseEmailInteractionCurrent, queue, parameters.TryGetValue<string, object>("Reason") as KeyValueCollection, parameters.TryGetValue<string, object>("Extensions") as KeyValueCollection, false, str1);

2) Email saving sequence:
Genesyslab.Desktop.Modules.OpenMedia.EmailModule:
      commandManager.AddCommandToChainOfCommand("InteractionEmailSave", (IList<CommandActivator>) new List<CommandActivator>()
      {
        new CommandActivator()
        {
          CommandType = typeof (SaveEmailCommand),
          Name = "Save"
        }
      });

Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email.SaveEmailCommand:
      this.eMailService.Update(interactionEmail.EntrepriseEmailInteractionCurrent);

CONCLUSION: Before sending reply email, if interactionEmail.EntrepriseEmailInteractionCurrent.IdType.Subtype != "OutboundNew":
1) attachments must be added from the parent interaction.
2) the interaction must be saved using the "InteractionEmailSave" command chain with reply interaction as a parameter.

-		EntrepriseEmailInteractionCurrent	{01VJ4TJRJ72M403R}	Genesyslab.Enterprise.Model.Interaction.IEmailInteraction {Genesyslab.Enterprise.Interaction.EmailInteraction}
		InteractionSubtype	"OutboundReply"	string

!!! To handle interaction content saving from the contact directory history tab apply the following changes:
Genesyslab.Desktop.Modules.OpenMedia.dll
namespace Genesyslab.Desktop.Modules.OpenMedia.Windows.Interactions.MediaView
internal class InteractionInboundEmailControlExtension : InteractionControlExtension, IInteractionControlExtension
public IList<RequestedAction> RequestActions(string capacity, ActionTarget target, object context)
case ActionTarget.InteractionActionFromContactHistory:
!!! add at the end of case statement this code:
		try
		{
			IWInteraction iwinteraction = contextDictionary.TryGetValue("interaction") as IWInteraction;
			list.Add(new RequestedAction
			{
				Type = text4,
				NameId = "Common.Images.Interaction.Email.Save",
				TooltipId = "Common.Images.Interaction.Email.Save",
				ImageId = "Common.Images.Interaction.Email.Save",
				Order = 50 + list.Count,
				IsPossible = flag,
				IsDefault = true,
				Action = delegate
				{
					this.SaveAttachments(iwinteraction);
				}
			});
			return list;
		}
		catch (Exception ex3)
		{
			this.log.Error("Adding Attachments action exception: " + ex3.ToString());
			return list;
		}

!!! add this function to the class:

private void SaveAttachments(object interacion)
{
	string value = "";
	if (interacion is IWInteraction)
	{
		value = (interacion as IWInteraction).Id;
	}
	else if (interacion is WorkbinInteraction)
	{
		value = (interacion as WorkbinInteraction).Id;
	}
	Utils.ExecuteAsynchronousCommand(this.commandManager.GetChainOfCommandByName("SaveAttachments"), new Dictionary<string, object>
	{
		{
			"CommandParameter",
			value
		}
	}, null);
}
