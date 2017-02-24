using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Threading;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email;
using System.IO;
using System.Text;
using System.Linq;
using Genesyslab.Platform.Commons.Protocols;
using Genesyslab.Platform.Commons.Logging;
using Genesyslab.Platform.Contacts.Protocols;
using Genesyslab.Platform.Contacts.Protocols.ContactServer;
using Genesyslab.Platform.Contacts.Protocols.ContactServer.Requests;
using Genesyslab.Platform.Contacts.Protocols.ContactServer.Events;
using Genesyslab.Desktop.Modules.Core.SDK.Configurations;
using System.Net.Mail;
using System.Reflection;

namespace Adventus.Modules.Email
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
					//var cfg = container.Resolve<IConfigurationService>();
					//var statServer = cfg.AvailableConnections.Where(item => item.Type.Equals(Genesyslab.Platform.Configuration.Protocols.Types.CfgAppType.CFGStatServer)).ToList();
					//var ServerInformation = statServer[0].ConnectionParameters[0].ServerInformation;

					// connection to contact server
					UniversalContactServerProtocol ucsConnection;
#if FOR_EE
					ucsConnection = new UniversalContactServerProtocol(new Endpoint("UniversalContactServer", "ling.lauteri.inter", 5130));
#else
					ucsConnection = new UniversalContactServerProtocol(new Endpoint("eS_UniversalContactServer", "genesys1", 6120));
#endif

					//ucsConnection = new UniversalContactServerProtocol(new Endpoint(statServer[0].Name, ServerInformation.Host.Name, Int32.Parse(ServerInformation.Port)));
					ucsConnection.Opened += new EventHandler(ucsConnection_Opened);
					ucsConnection.Error += new EventHandler(ucsConnection_Error);
					ucsConnection.Closed += new EventHandler(ucsConnection_Closed);
					try
					{
						ucsConnection.Open();
					}
					catch (Exception e)
					{
						MessageBox.Show(string.Format("Connection to UniversalContactServer failed. Save operation stopped. {0}", e.ToString()), "Attention");
						ucsConnection.Opened -= new EventHandler(ucsConnection_Opened);
						ucsConnection.Error -= new EventHandler(ucsConnection_Error);
						ucsConnection.Closed -= new EventHandler(ucsConnection_Closed);
						return false;
					}

					// download attachments and store in filesystem
					//IList<IAttachmentGraphic> attachments = new List<IAttachmentGraphic>();

					RequestGetInteractionContent request = new RequestGetInteractionContent();
					request.InteractionId = interaction.EntrepriseInteractionCurrent.Id;
					request.IncludeBinaryContent = true;
					request.IncludeAttachments = true;

					EventGetInteractionContent eventGetIxnContent = (EventGetInteractionContent)ucsConnection.Request(request);

					String subject = eventGetIxnContent.InteractionAttributes.Subject;
					//String key = eventGetIxnContent.InteractionAttributes.Id;
					AttachmentList attachmentList = eventGetIxnContent.Attachments;
					InteractionContent interactionContent = eventGetIxnContent.InteractionContent;

					// get attachment names
					List<string> attachmentNames = new List<string>();  // for adding to message body
					if (attachmentList != null)
					{
						foreach (Genesyslab.Platform.Contacts.Protocols.ContactServer.Attachment attachment in attachmentList)
						{
							attachmentNames.Add(attachment.TheName);
						}

						// check for attachments with the same name
						List<string> duplicates = attachmentNames.GroupBy(x => x.ToLower())
							.Where(x => x.Count() > 1)
							.Select(x => x.Key)
							.ToList();

						attachmentNames.Clear();

						foreach (Genesyslab.Platform.Contacts.Protocols.ContactServer.Attachment attachment in attachmentList)
						{
							//IAttachmentGraphic item = this.container.Resolve<IAttachmentGraphic>();
							//item.DocumentId = attachment.Id;
							//item.DataSourceType = DataSourceType.Main;
							string documentName = attachment.TheName;
							if (duplicates.Contains(documentName, StringComparer.OrdinalIgnoreCase))
							{
								documentName = attachment.DocumentId + "_" + documentName;
							}

							attachmentNames.Add(documentName); // for adding to message body
							DownloadAttachment(ucsConnection, attachment.DocumentId, documentName, subject);  // Saves Attachment on disk. Document path also saved in Model.EmailPartsPath
						}
					}

					if (ucsConnection.State != ChannelState.Closed && ucsConnection.State != ChannelState.Closing)
					{
						ucsConnection.Close();
						ucsConnection.Dispose();
					}

					string messageFrom = (interaction.GetAttachedData("FromAddress") ?? "").ToString();
					string messageTo = interactionEmail.EntrepriseEmailInteractionCurrent.To[0];
					string messageDate = interactionEmail.EntrepriseEmailInteractionCurrent.StartDate.ToString("dd.MM.yyyy HH:mm");

					string messageText = interactionEmail.EntrepriseEmailInteractionCurrent.MessageText;    /**< without html formatting */
					string structuredMessageText = interactionEmail.EntrepriseEmailInteractionCurrent.StructuredText;   /**< with html formatting */
					// mangle email subject
					string subj = RemoveSpecialChars(interactionEmail.EntrepriseEmailInteractionCurrent.Subject ?? "");
					if (subj.Length > MAX_SUBJECT_LENGTH) subj = subj.Substring(0, MAX_SUBJECT_LENGTH);

					string outputFolder = GetOutputFolder(subj);
					string path = Path.Combine(outputFolder, subj + ".eml");
					if (!Directory.Exists(outputFolder))
					{
						Directory.CreateDirectory(outputFolder);
					}

					// Binary content available. This is for incoming emails. They already have traveled the Business Process (URS).
					if (interactionContent.Content != null)
					{
						if (!Model.EmailPartsInfoStored) Model.EmailPartsPath.Add(path);
						if (File.Exists(path))   /**< file is already on disk */
						{
							return false;  //return and continue command chain
						}
						else
						{
							File.WriteAllBytes(path, interactionContent.Content);
						}
					}
					// Binary content not available. It is for outgoing emails, since email composed by agent may not be the final version sent out by Genesys. It can be changed by Business Process (URS).
					// Here we save email created by agent.
					else
					{
						MailMessage mailMessage = new MailMessage(messageFrom, messageTo);
						mailMessage.Subject = interactionEmail.EntrepriseEmailInteractionCurrent.Subject ?? "";
						if (structuredMessageText != null)
						{
							mailMessage.Body = structuredMessageText;
							mailMessage.IsBodyHtml = true;
						}
						else
						{
							mailMessage.Body = messageText;
							mailMessage.IsBodyHtml = false;
						}
						// Add attachments
						foreach (string pathToAttachment in Model.EmailPartsPath)
						{
							mailMessage.Attachments.Add(new System.Net.Mail.Attachment(pathToAttachment));
						}

						if (!Model.EmailPartsInfoStored) Model.EmailPartsPath.Add(path);
						if (File.Exists(path))   /**< don't download attachment if it's already on disk */
						{
							return false;	//return and continue command chain
						}
						if (mailMessage != null)
						{
							mailMessage.Save(path);
						}
						mailMessage.Dispose();
					}
					// The following files must be on disk: for incoming email - .eml + attachments; for outgoing - only .eml
					// delete all files except .eml for outgoing email

					if(interactionEmail.EntrepriseEmailInteractionCurrent.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.Out)
					{
						foreach (string p in Model.EmailPartsPath)
						{
							try
							{
								if (!p.EndsWith("eml"))
								{
									File.Delete(p);
								}
							}
							catch (Exception ex)
							{
								MessageBox.Show(string.Format("Cannot delete file {0}: {1}", p, ex.ToString()), "Attention");
							}
						}
						Model.EmailPartsPath.RemoveAll(x => !x.EndsWith("eml"));
					}

					//if (messageText != null && messageText.Length != 0)     // for plain text messages
					//{
					//	StringBuilder messageTextModified = new StringBuilder();
					//	messageTextModified.AppendLine("Subject: " + messageSubject);
					//	messageTextModified.AppendLine("From: " + messageFrom);
					//	messageTextModified.AppendLine("Date: " + messageDate);
					//	messageTextModified.AppendLine("To: " + messageTo);
					//	messageTextModified.AppendLine(string.Empty);
					//	if (attachmentNames.Count > 0)
					//	{
					//		messageTextModified.AppendLine(String.Format("{0} attachments:", attachmentNames.Count));
					//		foreach (string attachmentName in attachmentNames)
					//		{
					//			messageTextModified.AppendLine(attachmentName);
					//		}
					//		messageTextModified.AppendLine(string.Empty);
					//	}
					//	messageTextModified.Append(messageText);
					//	SaveMessage(messageTextModified.ToString(), false, false);
					//}
					//if (structuredMessageText != null && structuredMessageText.Length != 0)     // for html messages
					//{
					//	StringBuilder messageTextModified = new StringBuilder();
					//	messageTextModified.AppendLine("<!DOCTYPE html><html><head><title>" + messageSubject + "</title></head><body><p>");
					//	messageTextModified.AppendLine("<b>Subject: </b>" + messageSubject + "<br>");
					//	messageTextModified.AppendLine("<b>From: </b>" + messageFrom + "<br>");
					//	messageTextModified.AppendLine("<b>Date: </b>" + messageDate + "<br>");
					//	messageTextModified.AppendLine("<b>To: </b>" + messageTo + "<br><br>");
					//	if (attachmentNames.Count > 0)
					//	{
					//		messageTextModified.AppendLine(String.Format("<b>{0} attachments:</b><br>", attachmentNames.Count));
					//		foreach (string attachmentName in attachmentNames)
					//		{
					//			messageTextModified.AppendLine("<a href=\"" + attachmentName + "\">" + attachmentName + "</a><br>");
					//		}
					//		messageTextModified.AppendLine("<br>");
					//	}
					//	SaveMessage(structuredMessageText, true, true);  // save original email body in email.html file
					//	messageTextModified.AppendLine("<a href=\"original_email.html\">Original email</a>");
					//	messageTextModified.AppendLine("</body></html>");
					//	SaveMessage(messageTextModified.ToString(), true, false);  // save modified email body in subject.html file
					//}

					Model.EmailPartsInfoStored = true;
					StringBuilder sb = new StringBuilder();
					sb.Append("Interaction ID: " + interaction.EntrepriseInteractionCurrent.Id);
					sb.AppendLine(Environment.NewLine);
					sb.Append("From: " + interaction.GetAttachedData("FromAddress"));
					sb.AppendLine(Environment.NewLine);
					//sb.Append("Subject: " + interaction.GetAttachedData("Subject"));
					sb.Append("Subject: " + interactionEmail.EntrepriseEmailInteractionCurrent.Subject ?? "");
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

		private string GetOutputFolder(string subj)
		{
            string defaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
			return string.Format(@"{0}\{1}", defaultDirectory, subj);
		}

		private void ucsConnection_Closed(object sender, EventArgs e)
		{
			//throw new NotImplementedException();
		}

		private void ucsConnection_Error(object sender, EventArgs e)
		{
			MessageBox.Show(string.Format("Error in connection to UniversalContactServer. {0}", e.ToString()), "Attention");
		}

		private void ucsConnection_Opened(object sender, EventArgs e)
		{
			//throw new NotImplementedException();
		}

		delegate bool ExecuteDelegate(IDictionary<string, object> parameters, IProgressUpdater progressUpdater);

/** \brief Downloads email attachments and stores them to disk
 *  \param attachmentGraphic contains data about attachment
 *  \return full path of attachment on disk
 */       
        public string DownloadAttachment(UniversalContactServerProtocol ucsConnection, string documentId, string documentName, string emailSubject)
        {
            try
			{
                string defaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
				string subj = RemoveSpecialChars(emailSubject);

				if (subj.Length > MAX_SUBJECT_LENGTH) subj = subj.Substring(0, MAX_SUBJECT_LENGTH);
				string str = string.Format(@"{0}\{1}", defaultDirectory, subj);
				string path = Path.Combine(str, documentName);
				if (!Model.EmailPartsInfoStored) Model.EmailPartsPath.Add(path);

				if (File.Exists(path))   /**< don't download attachment if it's already on disk */
				{
					return path;
				}
				if (!Directory.Exists(str))
				{
					Directory.CreateDirectory(str);
				}

				RequestGetDocument request = new RequestGetDocument();
				request.DocumentId = documentId;
				request.IncludeBinaryContent = true;
				EventGetDocument eventGetDoc = (EventGetDocument)ucsConnection.Request(request);
				File.WriteAllBytes(path, eventGetDoc.Content);
			}
			catch (Exception exception)
            {
				MessageBox.Show(string.Format("Exception in DownloadAttachment. {0}", exception.ToString()), "Attention");
            }
            return null;
        }

		/** \brief removes special and whitespace chars from filename 
		 *  \param src source string
		 */
		private string RemoveSpecialChars(string src)
		{
			char[] chars = {'<', '>', ':', '\"', '\'', '/','\\', '|', '?', '*'}; // these chars are not allowed in filenames and path names
			var sb = new StringBuilder();
			for (int i = 0; i < src.Length; i++)
			{
		        if (!chars.Contains(src[i]))
					sb.Append(src[i]);
			}
			string res = sb.ToString();
			res = (res.Length > 0 && (!string.IsNullOrWhiteSpace(res))) ? res : "Empty Subject";
			return res;
		}

/** \brief Saves email message to disk.
 *  \param message contains the email body without attachments
 *  \param isStructured true if saving message with (html) formatting (IInteractionEmail.EntrepriseEmailInteractionCurrent.StructuredText),
 *           false if saving plain text message (IInteractionEmail.EntrepriseEmailInteractionCurrent.MessageText)
 *  \param originalHtmlFile true if saving email in file email.html, false if saving in subject.html file. Subject.html includes <a href='original_email.html'></a>.
 *  \return full path of message on disk
 */
    //    public string SaveMessage(string message, bool isStructured, bool originalHtmlFile)
    //    {
    //        try
    //        {
    //            string defaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
				//string subj = RemoveSpecialChars(interactionEmail.EntrepriseEmailInteractionCurrent.Subject);

				//if (subj.Length > MAX_SUBJECT_LENGTH) subj = subj.Substring(0,MAX_SUBJECT_LENGTH);
    //            string str = string.Format(@"{0}\{1}", defaultDirectory, subj);
    //            string path;
    //            if(!isStructured)
    //                path = Path.Combine(str, subj + ".txt");
    //            else
    //                path = Path.Combine(str, originalHtmlFile ? "original_email.html" : subj + ".html");

    //            if (!Model.EmailPartsInfoStored) Model.EmailPartsPath.Add(path);

    //            if (File.Exists(path))   /**< don't download attachment if it's already on disk */
    //            {
    //                return path;
    //            }
    //            if (!Directory.Exists(str))
    //            {
    //                Directory.CreateDirectory(str);
    //            }
    //            if (message != null)
    //            {
    //                FileStream stream = new FileInfo(path).Open(FileMode.Create, FileAccess.Write);
    //                using (StreamWriter s = new StreamWriter(stream))
    //                    s.Write(message);
    //                stream.Close();
    //                return path;
    //            }
    //            else
    //            {
    //                MessageBox.Show(string.Format("Email does not contain body. {0}", path), "Attention");
    //                return null;
    //            }
    //        }
    //        catch (Exception exception)
    //        {
    //            MessageBox.Show(string.Format("Exception in SaveMessage. {0}", exception.ToString()), "Attention");
    //        }
    //        return null;
    //    }
    }

	public static class MailMessageExt
{
    public static void Save(this MailMessage Message, string FileName)
    {
        Assembly assembly = typeof(SmtpClient).Assembly;
        Type _mailWriterType = 
          assembly.GetType("System.Net.Mail.MailWriter");

        using (FileStream _fileStream = 
               new FileStream(FileName, FileMode.Create))
        {
            // Get reflection info for MailWriter contructor
            ConstructorInfo _mailWriterContructor =
                _mailWriterType.GetConstructor(
                    BindingFlags.Instance | BindingFlags.NonPublic,
                    null,
                    new Type[] { typeof(Stream) }, 
                    null);

            // Construct MailWriter object with our FileStream
            object _mailWriter = 
              _mailWriterContructor.Invoke(new object[] { _fileStream });

            // Get reflection info for Send() method on MailMessage
            MethodInfo _sendMethod =
                typeof(MailMessage).GetMethod(
                    "Send",
                    BindingFlags.Instance | BindingFlags.NonPublic);

            // Call method passing in MailWriter
			_sendMethod.Invoke(
				Message,
				BindingFlags.Instance | BindingFlags.NonPublic,
				null,
				new object[] { _mailWriter, true, true },
				null);

            // Finally get reflection info for Close() method on our MailWriter
            MethodInfo _closeMethod =
                _mailWriter.GetType().GetMethod(
                    "Close",
                    BindingFlags.Instance | BindingFlags.NonPublic);

            // Call close method
            _closeMethod.Invoke(
                _mailWriter,
                BindingFlags.Instance | BindingFlags.NonPublic,
                null,
                new object[] { },
                null);
        }
    }
}
}
