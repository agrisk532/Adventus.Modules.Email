return this.Interaction.ViewData["OutboundEmailView_Attachments"] as ObservableCollection<IAttachmentGraphic>;

interaction.ViewData["OutboundEmailView_Attachments"]
UCS port: 6120
ucsConnection = new UniversalContactServerProtocol(new Endpoint("eS_UniversalContactServer", "genesys1", 6120));

eml format works with .net ver 4.5. That is due to MailMessage.Send() function.