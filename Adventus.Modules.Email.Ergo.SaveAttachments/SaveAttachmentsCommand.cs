using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Threading;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email;
using Genesyslab.Desktop.Modules.Windows;
using Genesyslab.Enterprise.Model.Contact;
using Genesyslab.Desktop.Modules.Windows.Views.Common.AttachmentView;
using Genesyslab.Enterprise.Services;
using Genesyslab.Enterprise.Model.ServiceModel;
using Genesyslab.Enterprise.Model;
using System.IO;
using System.Text;
using System.Linq;
using Genesyslab.Desktop.Modules.Core.SDK.Protocol;
using Genesyslab.Enterprise.Model.Channel;
using Genesyslab.Platform.Commons.Protocols;
using Genesyslab.Enterprise.Commons.Collections;
using Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions;
using Microsoft.Practices.Unity;
using Genesyslab.Platform.Commons.Logging;
using Genesyslab.Desktop.Infrastructure.ExceptionAnalyze;
using Genesyslab.Desktop.Infrastructure.Events;

namespace Adventus.Modules.Email.Ergo.SaveAttachments
{
/** \class SaveAttachmentsCommand
 *  \brief Saves email attachments to disk
 */
    class SaveAttachmentsCommand : IElementOfCommand
    {
        readonly IObjectContainer container;
        readonly ILogger log;
        public string Name { get; set; }
        public SaveAttachmentsViewModel Model { get; set; }
        public IInteraction interaction { get; set; }
        public IInteractionEmail interactionEmail { get; set; }
		public const int MAX_SUBJECT_LENGTH = 20;
        
        public SaveAttachmentsCommand(IObjectContainer container, ILogger logger)
        {
          this.container = container;
          this.log = logger;
          log.Info("SaveAttachmentsCommand() entered");

        }

/** \brief Command implementation
 *  \param parameters - input data
 *  \param progress
 *  \return true to stop execution of the command chain; otherwise, false.
 */       
        public bool Execute(IDictionary<string, object> parameters, IProgressUpdater progress)
        {
          // To go to the main thread
            if (Application.Current.Dispatcher != null && !Application.Current.Dispatcher.CheckAccess())
            {
                object result = Application.Current.Dispatcher.Invoke(DispatcherPriority.Send, new ExecuteDelegate(Execute), parameters, progress);
                return (bool)result;
            }
            else
            {
                Model = parameters["Model"] as SaveAttachmentsViewModel;
                interaction = Model.Interaction;
                interactionEmail = interaction as IInteractionEmail;

                if (interaction == null)
                {
                    MessageBox.Show("Interaction is NULL");
                    return true;
                }
                else
                if(interactionEmail == null)
                {
                    MessageBox.Show("Interaction is not of IInteractionEmail type");
                    return true;  // stop execution of command chain
                }
                else
                {
// save the message body
                    string messageText = interactionEmail.EntrepriseEmailInteractionCurrent.MessageText;    /**< without html formatting */
                    string structuredMessageText = interactionEmail.EntrepriseEmailInteractionCurrent.StructuredText;   /**< with html formatting */
                    if (messageText != null && messageText.Length != 0) 
                        SaveMessage(messageText, false);
                    if (structuredMessageText != null && structuredMessageText.Length != 0) 
                        SaveMessage(structuredMessageText, true);

                    if ((interactionEmail.EntrepriseEmailAttachments != null) && (interactionEmail.EntrepriseEmailAttachments.Count > 0))
                    {
                        ICollection<IAttachment> attachments = new List<IAttachment>();
                        IContactService service = this.container.Resolve<IEnterpriseServiceProvider>().Resolve<IContactService>("contactService");
// get attachment names only
                        attachments = service.GetAttachments(interaction.EntrepriseInteractionCurrent);
// check for attachments with the same name
                        List<string> duplicates = attachments.GroupBy(x => x.Name.ToLower())
                            .Where(x => x.Count() > 1)
                            .Select(x => x.Key)
                            .ToList();

                        foreach (IAttachment attachment in attachments)
                        {
                            if (attachment != null)
                            {
                                IAttachmentGraphic item = this.container.Resolve<IAttachmentGraphic>();
                                item.DocumentId = attachment.Id;
                                item.DataSourceType = DataSourceType.Main;
                                if (duplicates.Contains(attachment.Name, StringComparer.OrdinalIgnoreCase))
                                {
                                    item.DocumentName = attachment.Id + "_" + attachment.Name;
                                }
                                else
                                {
                                    item.DocumentName = attachment.Name;
                                }
                                item.DocumentSize = attachment.Size.ToString();
                                DownloadAttachment(item, interaction);
                            }
                        }
                    }

//                        KeyValueCollection attachedData = new KeyValueCollection();
//                        attachedData = interaction.GetAllAttachedData();

                    Model.EmailPartsInfoStored = true;
                    StringBuilder sb = new StringBuilder();
                    sb.Append("Interaction ID: " + interaction.EntrepriseInteractionCurrent.Id);
                    sb.AppendLine(Environment.NewLine);
                    sb.Append("From: " + interaction.GetAttachedData("FromAddress"));
                    sb.AppendLine(Environment.NewLine);
                    sb.Append("Subject: " + interaction.GetAttachedData("Subject"));
                    sb.AppendLine(Environment.NewLine);
                    sb.Append("Email parts: ");
                    sb.AppendLine(Environment.NewLine);
                    for (int i = 0; i < Model.EmailPartsPath.Count; i++)
                    {
                        sb.AppendLine(Model.EmailPartsPath[i]);
                        sb.AppendLine();
                    }
                    MessageBox.Show(sb.ToString(), "Information");

                    return false;
                }
            }
        }

        delegate bool ExecuteDelegate(IDictionary<string, object> parameters, IProgressUpdater progressUpdater);

/** \brief Downloads email attachments and stores them to disk
 *  \param attachmentGraphic contains data about attachment
 *  \return full path of attachment on disk
 */       
        public string DownloadAttachment(IAttachmentGraphic attachmentGraphic, IInteraction interaction)
        {
            try
            {

				string defaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
				//string subj = (interaction.GetAttachedData("Subject") ?? "Empty Subject").ToString();
				string subj = interactionEmail.EntrepriseEmailInteractionCurrent.Subject ?? "Empty Subject";
				if (subj.Length > MAX_SUBJECT_LENGTH) subj = subj.Substring(0,MAX_SUBJECT_LENGTH);
                string str = string.Format(@"{0}\{1}", defaultDirectory, subj);
                attachmentGraphic.DirectoryFullName = str;
                string path = Path.Combine(str, attachmentGraphic.GetValidFileName());
                if(!Model.EmailPartsInfoStored) Model.EmailPartsPath.Add(path);

                if (File.Exists(path))   /**< don't download attachment if it's already on disk */
                {
                    return path;
                }
                if (!Directory.Exists(str))
                {
                    Directory.CreateDirectory(str);
                }
                IAttachment attachment = this.LoadAttachment(attachmentGraphic.DocumentId, attachmentGraphic.DataSourceType);
                if (attachment != null)
                {
                    FileStream stream = new FileInfo(path).Open(FileMode.Create, FileAccess.Write);
                    stream.Write(attachment.Content, 0, attachment.Size.Value);
                    stream.Close();
                    return path;
                }
                else
                {
                    MessageBox.Show("DownloadAttachment is null. Can't downloaded the file {0}", path );
                    return null;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Exception in DownloadAttachment. {0}", exception.ToString());
            }
            return null;
        }

/** \brief Gets email attachments from the contact server
 *  \param attachmentId global Id of attachment
 *  \param dataSourceType Main or Archive
 *  \return email attachment
 */       
        private IAttachment LoadAttachment(string attachmentId, DataSourceType dataSourceType)
        {
            IContactService service = this.container.Resolve<IEnterpriseServiceProvider>().Resolve<IContactService>("contactService");
            Genesyslab.Desktop.Modules.Core.SDK.Contact.IContactService service2 = this.container.Resolve<Genesyslab.Desktop.Modules.Core.SDK.Contact.IContactService>();
            Genesyslab.Enterprise.Model.Channel.IClientChannel channel = this.container.Resolve<Genesyslab.Desktop.Modules.Core.SDK.Protocol.IChannelManager>().Register(service2.UCSApp, "IW@ContactService");
            if ((channel != null) && (channel.State == ChannelState.Opened))
            {
                return service.GetAttachment(channel, attachmentId, dataSourceType, WindowsOptions.Default.Emailattachmentdownloadtimeout);
            }
            return null;
        }

/** \brief Saves email message to disk.
 *  \param message contains the email body without attachments
 *  \param isStructured true if saving message with (html) formatting (IInteractionEmail.EntrepriseEmailInteractionCurrent.StructuredText),
 *           false if saving plain text message (IInteractionEmail.EntrepriseEmailInteractionCurrent.MessageText)
 *  \return full path of message on disk
 */
        public string SaveMessage(string message, bool isStructured)
        {
            try
            {
                string defaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
				//string subj = (interaction.GetAttachedData("Subject") ?? "Empty Subject").ToString();
				string subj = interactionEmail.EntrepriseEmailInteractionCurrent.Subject ?? "Empty Subject";
				if (subj.Length > MAX_SUBJECT_LENGTH) subj = subj.Substring(0,MAX_SUBJECT_LENGTH);
                string str = string.Format(@"{0}\{1}", defaultDirectory, subj);
                string path;
                if(!isStructured)
                    path = Path.Combine(str, "email.txt");
                else
                    path = Path.Combine(str, "email.html");
                if (!Model.EmailPartsInfoStored) Model.EmailPartsPath.Add(path);

                if (File.Exists(path))   /**< don't download attachment if it's already on disk */
                {
                    return path;
                }
                if (!Directory.Exists(str))
                {
                    Directory.CreateDirectory(str);
                }
                if (message != null)
                {
                    FileStream stream = new FileInfo(path).Open(FileMode.Create, FileAccess.Write);
                    using (StreamWriter s = new StreamWriter(stream))
                        s.Write(message);
                    stream.Close();
                    return path;
                }
                else
                {
                    MessageBox.Show("Email does not contain body. {0}", path);
                    return null;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Exception in SaveMessage. {0}", exception.ToString());
            }
            return null;
        }
    }
}
