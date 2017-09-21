using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using System.Collections.Generic;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email;
using Genesyslab.Platform.Commons.Protocols;
using Genesyslab.Platform.Commons.Logging;
using Genesyslab.Platform.Contacts.Protocols;
using Genesyslab.Platform.Contacts.Protocols.ContactServer;
using Genesyslab.Platform.Contacts.Protocols.ContactServer.Requests;
using Genesyslab.Platform.Contacts.Protocols.ContactServer.Events;
using Genesyslab.Desktop.Modules.Core.SDK.Configurations;
using Genesyslab.Platform.Commons.Collections;
using Genesyslab.Platform.ApplicationBlocks.ConfigurationObjectModel.CfgObjects;
using Genesyslab.Platform.ApplicationBlocks.ConfigurationObjectModel.Queries;
using Genesyslab.Enterprise.Model.ServiceModel;
using Genesyslab.Enterprise.Services;
using Genesyslab.Desktop.Modules.Core.Model.Agents;
using MimeKit;

namespace Adventus.Modules.Email
{
/** \class SaveAttachmentsCommand
 *  \brief Saves email attachments to disk
 */
    class SaveAttachmentsCommand : IElementOfCommand
    {
        readonly IObjectContainer container;
        readonly ILogger log;
        public string Name { get; set; }  // to comply with IElementOfCommand
		public int subjectLength;
		public string OutputFolderName;		// for saving .eml and attachments
		readonly IConfigurationService configurationService;
		public SaveAttachmentsViewModelBase Model;
		public Genesyslab.Enterprise.Model.Interaction.IEmailInteraction enterpriseEmailInteraction;
		
		// for calls from ContactDirectory History tab
		public bool isCalledFromHistory;

		public Genesyslab.Enterprise.Services.IContactService service;
		public Genesyslab.Desktop.Modules.Core.SDK.Contact.IContactService service2;
		public Genesyslab.Enterprise.Model.Channel.IClientChannel channel;

		// Config server parameters for user and application. Email and attachment save path
		private const string CONFIG_SECTION_NAME_EMAIL_SAVE	= "custom-email-content-save";		// section name
		private const string CONFIG_OPTION_NAME_EMAIL_SAVE_PATH = "email-content-save-path";	// name of the folder where .eml and attachments will be stored
		// Config server parameters for user and application. Email save option name. It specifies what to save, email binary format, attachments or both
		private const string CONFIG_OPTION_NAME_INBOUND_EMAIL_SAVE_OPTION = "email-content-save-options-inbound";  // values: eml, attachments, all (eml+attachments)
		private const string CONFIG_OPTION_NAME_OUTBOUND_EMAIL_SAVE_OPTION = "email-content-save-options-outbound"; // values: eml, attachments, all (eml+attachments)
		// Config server parameter for subject truncation length
		private const string CONFIG_OPTION_NAME_EMAIL_SAVE_SUBJECT_LENGTH = "email-content-save-subject-length";
		private const string DEFAULT_SUBJECT_LENGTH = "25";		// default value for option CONFIG_OPTION_NAME_EMAIL_SAVE_SUBJECT_LENGTH
		private const string METHOD_NAME = "Adventus.Modules.Email.SaveAttachmentsCommand(): ";

		// constructor
        public SaveAttachmentsCommand(IObjectContainer container, ILogger logger)
        {
			this.container = container;
			this.configurationService = container.Resolve<IConfigurationService>();
			this.log = logger;
			log.Info(METHOD_NAME + "entered");
			OutputFolderName = String.Empty;
			isCalledFromHistory = false; // called from active online interaction
			service = container.Resolve<IEnterpriseServiceProvider>().Resolve<IContactService>("contactService");
			service2 = container.Resolve<Genesyslab.Desktop.Modules.Core.SDK.Contact.IContactService>();
			channel = container.Resolve<Genesyslab.Desktop.Modules.Core.SDK.Protocol.IChannelManager>().Register(service2.UCSApp, "IW@ContactService");
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
				try
				{
					if(parameters.ContainsKey("Model"))	
					{
						SaveAttachmentsViewModel m = parameters["Model"] as SaveAttachmentsViewModel;
						SaveAttachmentsViewModelH mH = parameters["Model"] as SaveAttachmentsViewModelH;

						if(m != null)	// ISaveAttachmentsViewModel used (active interaction)
						{
			                IInteraction interaction = m.Interaction;
			                if (interaction == null)
			                {
								ShowAndLogErrorMsg("Interaction is NULL. Email saving terminated.");
			                    return true;	// stop execution of command chain
			                }

			                IInteractionEmail interactionEmail = interaction as IInteractionEmail;
			                if(interactionEmail == null)
			                {
								ShowAndLogErrorMsg("Interaction is not of IInteractionEmail type. Email saving terminated.");
			                    return true;	// stop execution of command chain
			                }

							enterpriseEmailInteraction = interactionEmail.EntrepriseEmailInteractionCurrent;
							Model = m;
							isCalledFromHistory = false;
						}
						else
						if(mH != null)	// ISaveAttachmentsViewModelH used (interaction from history)
						{
							enterpriseEmailInteraction = service.GetInteractionContent(channel, mH.SelectedInteractionId, (Genesyslab.Enterprise.Services.DataSourceType)mH.Dst) as Genesyslab.Enterprise.Model.Interaction.IEmailInteraction;
							Model = mH;
							isCalledFromHistory = true;
						}
						else
						{
							ShowAndLogErrorMsg("Invalid view model type. Email saving terminated.");
							return true;	// stop execution of command chain
						}
					}
					else
					{
						ShowAndLogErrorMsg("Command parameter error. Email saving terminated.");
						return true;
					}
				}
				catch(Exception e)
				{
					ShowAndLogErrorMsg("Type error. Email saving terminated.");
					return true;
				}

					// get subjectLength
					subjectLength = GetSubjectLength();

					// Subfolder name for the output folder
					string SubjectTrimmed = GetSubjectTrimmed();

					// create folder where files will be written. It includes trimmed subject at the end of the path
					OutputFolderName = GetOutputFolderName(SubjectTrimmed);

					for (int i = 0; i < 3; i++)
					{
						if (i == 2)
						{
							ShowAndLogErrorMsg(String.Format("Cannot create output folder {0}. Email saving terminated.", OutputFolderName));
							return true; // Cannot create output folder. Stop execution of the command chain
						}

						try
						{
							if (!Directory.Exists(OutputFolderName))
							{
								Directory.CreateDirectory(OutputFolderName);
							}
							else
							{
								MessageBoxResult r = MessageBox.Show(String.Format("Folder \"{0}\" already exists. Overwrite?", OutputFolderName), "Question", MessageBoxButton.YesNo);
								if(r == MessageBoxResult.Yes)
								{
									Array.ForEach(Directory.GetFiles(OutputFolderName), File.Delete);
									break;
								}
								else
								if(r == MessageBoxResult.No)
								{
									return false;	// continue execution of the command chain
								}
							}
							break;
						}
						catch (Exception exception)
						{
							ShowAndLogErrorMsg(String.Format("Exception creating folder at {0}: {1}.", OutputFolderName, exception.Message));
							OutputFolderName = SetDesktopOutputFolder() + "\\" + SubjectTrimmed;
						}
					}

					// read contact server location from the config server
					var cfg = container.Resolve<IConfigurationService>();
					//var statServer = cfg.AvailableConnections.Where(item => item.Type.Equals(Genesyslab.Platform.Configuration.Protocols.Types.CfgAppType.CFGStatServer)).ToList();
					//var ServerInformation = statServer[0].ConnectionParameters[0].ServerInformation;
					var contactServer = cfg.AvailableConnections.Where(item => item.Type.Equals(Genesyslab.Platform.Configuration.Protocols.Types.CfgAppType.CFGContactServer)).ToList();
					var serverInformation = contactServer[0].ConnectionParameters[0].ServerInformation;
					var appName = contactServer[0].Name;
					var host = serverInformation.ConnectedHost.Name;
					var s_port = serverInformation.ConnectedPort;
					int port;
		            if (!Int32.TryParse(s_port, out port))
					{
						ShowAndLogErrorMsg(String.Format("Configured contact server port cannot be converted to int: {0}. Email not saved.", s_port));
						return true;	// Cannot parse server port. Stop execution of the command chain.
					}

					// connection to UCS contact server
					UniversalContactServerProtocol ucsConnection;
					//ucsConnection = new UniversalContactServerProtocol(new Endpoint("UniversalContactServer", "ling.lauteri.inter", 5130));
					//ucsConnection = new UniversalContactServerProtocol(new Endpoint("eS_UniversalContactServer", "genesys1", 6120));

					ucsConnection = new UniversalContactServerProtocol(new Endpoint(appName, host, port));

					ucsConnection.Opened += new EventHandler(ucsConnection_Opened);
					ucsConnection.Error += new EventHandler(ucsConnection_Error);
					ucsConnection.Closed += new EventHandler(ucsConnection_Closed);
					try
					{
						ucsConnection.Open();
					}
					catch (Exception e)
					{
						ShowAndLogErrorMsg(String.Format("Connection to UniversalContactServer failed. Email saving terminated: {0}", e.ToString()));
						ucsConnection.Opened -= new EventHandler(ucsConnection_Opened);
						ucsConnection.Error -= new EventHandler(ucsConnection_Error);
						ucsConnection.Closed -= new EventHandler(ucsConnection_Closed);
						return true;    // stop execution of command chain
					}

					RequestGetInteractionContent request = new RequestGetInteractionContent();
					request.InteractionId = enterpriseEmailInteraction.Id;
					request.IncludeBinaryContent = true;
					request.IncludeAttachments = true;
					request.DataSource = new NullableDataSourceType(Model.Dst);

					EventGetInteractionContent eventGetIxnContent = (EventGetInteractionContent)ucsConnection.Request(request);

					if (eventGetIxnContent == null)
					{
						ShowAndLogErrorMsg("Request to UniversalContactServer failed. Email saving terminated.");
						CloseUCSConnection(ucsConnection);
						return true;    // stop execution of the command chain
					}

					// create and save message

					//string messageFrom = (interaction.GetAttachedData("FromAddress") ?? "").ToString();

					string messageFrom = enterpriseEmailInteraction.From;
					//string messageFrom = "\"ERGO kahjukäsitlus\" <kahju@ergo.ee>";
					string messageTo;
					try
					{
						messageTo = enterpriseEmailInteraction.To[0] ?? "";
					}
					catch (Exception e)
					{
						MessageBox.Show("Please enter message recipient", "Attention");
						CloseUCSConnection(ucsConnection);
						return false;	 // continue execution of the command chain
					}
					if (String.IsNullOrEmpty(messageTo))
					{
						MessageBox.Show("Please enter message recipient", "Attention");
						CloseUCSConnection(ucsConnection);
						return false;	// continue execution of the command chain
					}
					string messageDate = enterpriseEmailInteraction.StartDate.ToString("dd.MM.yyyy HH:mm");
					string messageText = enterpriseEmailInteraction.MessageText;    /**< without html formatting */
					string structuredMessageText = enterpriseEmailInteraction.StructuredText;   /**< with html formatting */
					string emlFilePath = Path.Combine(OutputFolderName, SubjectTrimmed + ".eml");

					AttachmentList attachmentList = eventGetIxnContent.Attachments;
					InteractionContent interactionContent = eventGetIxnContent.InteractionContent;

					// read options from config server
					string opt = String.Empty;
					if (enterpriseEmailInteraction.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.In ||
						enterpriseEmailInteraction.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.Unknown)
					{
						opt = GetConfigurationOption(CONFIG_SECTION_NAME_EMAIL_SAVE, CONFIG_OPTION_NAME_INBOUND_EMAIL_SAVE_OPTION);
						if (opt == "eml")
						{
						// for inbound email binary content is available. We save it.
							SaveBinaryContent(messageFrom, messageTo, messageText, structuredMessageText, emlFilePath, interactionContent);
						}
						else
						if (opt == "attachments")
						{
							SaveAttachments(ucsConnection, attachmentList);
						}
						else
						{
							SaveAttachments(ucsConnection, attachmentList);
							SaveBinaryContent(messageFrom, messageTo, messageText, structuredMessageText, emlFilePath, interactionContent);
						}
					}
					else
					if (enterpriseEmailInteraction.IdType.Direction == Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.Out)
					{
						opt = GetConfigurationOption(CONFIG_SECTION_NAME_EMAIL_SAVE, CONFIG_OPTION_NAME_OUTBOUND_EMAIL_SAVE_OPTION);
						if (opt == "eml")
						{
						// for outbound emails binary content is not available. We must assemble it from agent created parts. This is not the email finally sent out by the business process.
						// for this option == eml, store attachments, assemble message including attachments, delete attachments
							SaveAttachments(ucsConnection, attachmentList);
							SaveBinaryContent(messageFrom, messageTo, messageText, structuredMessageText, emlFilePath, interactionContent);
							DeleteAttachments();
						}
						else
						if (opt == "attachments")
						{
							SaveAttachments(ucsConnection, attachmentList);
						}
						else
						{
							SaveAttachments(ucsConnection, attachmentList);
							SaveBinaryContent(messageFrom, messageTo, messageText, structuredMessageText, emlFilePath, interactionContent);
						}
					}
					else
					{
						ShowAndLogErrorMsg("I don't know what to do. It is neither inbound nor outbound email. Email saving terminated.");
						CloseUCSConnection(ucsConnection);
						return true;	// stop execution of command chain
					}

					CloseUCSConnection(ucsConnection);

					// show info in a messagebox
					Model.EmailPartsInfoStored = true;
					StringBuilder sb = new StringBuilder();
					sb.Append("Interaction ID: " + enterpriseEmailInteraction.Id);
					sb.AppendLine(Environment.NewLine);
					//sb.Append("From: " + interaction.GetAttachedData("FromAddress"));
					sb.Append("From: " + enterpriseEmailInteraction.From);
					sb.AppendLine(Environment.NewLine);
					//sb.Append("Subject: " + interaction.GetAttachedData("Subject"));
					sb.Append("Subject: " + enterpriseEmailInteraction.Subject ?? "");
					sb.AppendLine(Environment.NewLine);
					sb.Append("Email parts: ");
					sb.AppendLine(Environment.NewLine);
					for (int i = 0; i < Model.EmailPartsPath.Count; i++)
					{
						sb.AppendLine(Model.EmailPartsPath[i]);
						sb.AppendLine();
					}
					//MessageBox.Show(sb.ToString(), "Information");
					Model.Clear();
					return false;
			}
        }

		private int GetSubjectLength()
		{
			int sl;
			string s = GetConfigurationOption(CONFIG_SECTION_NAME_EMAIL_SAVE, CONFIG_OPTION_NAME_EMAIL_SAVE_SUBJECT_LENGTH) ?? DEFAULT_SUBJECT_LENGTH;
            if (!Int32.TryParse(s, out sl))
			{
				ShowAndLogInfoMsg(String.Format("Configured max Email subject length cannot be converted to int: {0}. Using {1} instead.", s, DEFAULT_SUBJECT_LENGTH));
				Int32.TryParse(DEFAULT_SUBJECT_LENGTH, out sl);	// Cannot parse server port. Stop execution of the command chain.
			}
			return sl;
		}

		private void DeleteAttachments()
		{
			// what about attachment filenames ending with .eml ?
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
					ShowAndLogErrorMsg(String.Format("Cannot delete file {0}: {1}", p, ex.ToString()));
				}
			}
			Model.EmailPartsPath.RemoveAll(x => !x.EndsWith("eml"));  
		}

		private void CloseUCSConnection(UniversalContactServerProtocol ucsConnection)
		{
			if (ucsConnection.State != ChannelState.Closed && ucsConnection.State != ChannelState.Closing)
			{
				ucsConnection.Opened -= new EventHandler(ucsConnection_Opened);
				ucsConnection.Error -= new EventHandler(ucsConnection_Error);
				ucsConnection.Closed -= new EventHandler(ucsConnection_Closed);
				ucsConnection.Close();
				ucsConnection.Dispose();
			}
		}

		private void SaveBinaryContent(string messageFrom, string messageTo, string messageText, string structuredMessageText, string emlFilePath, InteractionContent interactionContent)
		{
			// Binary content available. That is for incoming emails. They already have traveled the Business Process (URS).
			if (interactionContent.Content != null)
			{
				SaveEMLBinaryContent(interactionContent, emlFilePath);
			}
			// Binary content not available. That is for outgoing emails, since email composed by agent may not be the final version sent out by Genesys. It can be changed by Business Process (URS).
			// Here we save email created by agent.
			else
			{
				AssembleAndSaveEMLBinaryContent(messageFrom, messageTo, messageText, structuredMessageText, emlFilePath);
			}
		}

		// attachments must be saved on the file system before calling this method. Use method SaveAttachments(ucsConnection, attachmentList)
		private void AssembleAndSaveEMLBinaryContent(string messageFrom, string messageTo, string messageText, string structuredMessageText, string path)
		{
			var message = new MimeMessage();
			// get the name and address

			string name, addr;
			SplitAddress(messageFrom, out name, out addr);
			message.From.Add(new MailboxAddress(name, addr));
			SplitAddress(messageTo, out name, out addr);
			message.To.Add(new MailboxAddress(name, addr));
			HeaderList l = message.Headers;
			string s = enterpriseEmailInteraction.Subject ?? "";
			//l["Subject"] = @"=?utf-8?Q?" + Encoder.EncodeQuotedPrintable(s) + @"?=";
			// Base64 encoding
			//var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(s);
			//var subject = @"=?utf-8?B?" + System.Convert.ToBase64String(plainTextBytes) + @"?=";
			l["Subject"] = s;

			var builder = new BodyBuilder();

			// Set the plain-text version of the message text
			builder.TextBody = messageText;
			// Set the html version of the message text
			builder.HtmlBody = structuredMessageText;

			foreach (string pathToAttachment in Model.EmailPartsPath)
			{
				try
				{
					builder.Attachments.Add(pathToAttachment);
				}
				catch (Exception ex)
				{
					ShowAndLogErrorMsg(String.Format("Cannot attach file {0}: {1}", pathToAttachment, ex.ToString()));
				}
			}

			// Now we just need to set the message body and we're done
			message.Body = builder.ToMessageBody();
			try
			{
				message.WriteTo(path);
				if (!Model.EmailPartsInfoStored) Model.EmailPartsPath.Add(path);
			}
			catch (Exception ex)
			{
				ShowAndLogErrorMsg(String.Format("Cannot save file {0}: {1}", path, ex.ToString()));
			}

		}

		private static void SplitAddress(string address, out string name, out string addr)
		{
			int start = address.IndexOf("\"") + 1;
			int end;
			if(start != 0)
			{
				end = address.IndexOf("\"", start);
				name = address.Substring(start, end - start);
			}
			else
			{
				name = String.Empty;
			}

			start = address.IndexOf("<") + 1;
			if(start != 0)
			{
				end = address.IndexOf(">", start);
				addr = address.Substring(start, end - start);
			}
			else
			{
				addr = address;
			}
		}

		private void SaveEMLBinaryContent(InteractionContent interactionContent, string path)
		{
			try
			{
				File.WriteAllBytes(path, interactionContent.Content);
				if (!Model.EmailPartsInfoStored) Model.EmailPartsPath.Add(path);
			}
			catch (Exception ex)
			{
				ShowAndLogErrorMsg(String.Format("Cannot Save File {0}: {1}", path, ex.ToString()));
			}
		}

		private void SaveAttachments(UniversalContactServerProtocol ucsConnection, AttachmentList attachmentList)
		{
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
					string documentName = attachment.TheName;
					if (duplicates.Contains(documentName, StringComparer.OrdinalIgnoreCase))
					{
						documentName = attachment.DocumentId + "_" + documentName;
					}
					documentName = RemoveSpecialChars(documentName);
					//attachmentNames.Add(documentName); // for adding to message body
					DownloadAndSaveAttachment(ucsConnection, attachment.DocumentId, documentName);  // Saves Attachment on disk. Document path also saved in Model.EmailPartsPath
				}
			}
		}

		// Mangle Subject
		private string GetSubjectTrimmed()
		{
			string subj = RemoveSpecialChars(enterpriseEmailInteraction.Subject ?? "Empty subject");
			if (subj.Length > subjectLength) subj = subj.Substring(0, subjectLength);
			return subj;
		}

		// read configuration options
		// first read application level configuration option, if not set, read user level config option, if not set use default option
		private string GetConfigurationOption(string section, string option)
		{
			string opt = String.Empty;

			while(true)
			{
// user level configuration option
				try
				{
					Genesyslab.Platform.ApplicationBlocks.ConfigurationObjectModel.CfgObjects.CfgPerson cp = container.Resolve<IAgent>().ConfPerson;  // <--- TEST THIS
					Genesyslab.Platform.Commons.Collections.KeyValueCollection kvc = cp.UserProperties;
					Genesyslab.Platform.Commons.Collections.KeyValueCollection sect = (Genesyslab.Platform.Commons.Collections.KeyValueCollection) kvc[section];
					opt = (string)sect[option];
					opt = Environment.ExpandEnvironmentVariables(opt);
					break;
				}
		        catch (Exception ex)
		        {
					// fall through to the application options
					log.Info(String.Format(METHOD_NAME + "User level configuration option {0} not defined. Trying to read application level option.", option));
		        }

// application level configuration option
				try
				{
					//string name = configurationService.MyApplication.Name;
					CfgApplication app = configurationService.RetrieveObject<CfgApplication>((ICfgQuery)new CfgApplicationQuery()
					{
						Name = configurationService.MyApplication.Name
					});
	
					KeyValueCollection kvc = app.Options;
					KeyValueCollection sect = (KeyValueCollection) kvc[section];
					opt = (string)sect[option];
					opt = Environment.ExpandEnvironmentVariables(opt);
					break;
				}
				catch (Exception e)
				{
					opt = null;
					log.Info(String.Format(METHOD_NAME + "Exception reading application level configuration option {0}", option));
				}
	
				if(String.IsNullOrEmpty(opt))
				{
					opt = null;
					log.Info(String.Format(METHOD_NAME + "Configuration option {0} not defined", option));
				}
				break;
			}
			return opt;
		}

		// The folder where files will be written. If not configured, it is agent's PC Desktop folder
		private string GetOutputFolderName(string s)
		{
			string defaultDirectory = String.Empty;
			try
			{
				defaultDirectory = GetConfigurationOption(CONFIG_SECTION_NAME_EMAIL_SAVE, CONFIG_OPTION_NAME_EMAIL_SAVE_PATH);
			}
			catch (Exception e)
			{
				defaultDirectory = SetDesktopOutputFolder();
				log.Info(string.Format(METHOD_NAME + "Output folder configuration option not defined. Using a Desktop folder."));
			}

			if(String.IsNullOrEmpty(defaultDirectory))
			{
				defaultDirectory = SetDesktopOutputFolder();
				log.Info(string.Format(METHOD_NAME + "Output folder configuration option not defined. Using a folder on Desktop."));
			}

			return string.Format(@"{0}\{1}", defaultDirectory, s);
		}
//
		private string SetDesktopOutputFolder()
		{
			string defaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
			//MessageBox.Show(string.Format("Output folder " + CONFIG_SECTION_NAME_EMAIL_SAVE + "/" + CONFIG_OPTION_NAME_EMAIL_SAVE_PATH + " not configured.\nUsing {0}.", defaultDirectory), "Attention");
			return defaultDirectory;

		}
//
		private void ucsConnection_Closed(object sender, EventArgs e)
		{
			//throw new NotImplementedException();
		}

		private void ucsConnection_Error(object sender, EventArgs e)
		{
			ShowAndLogErrorMsg(String.Format("Error in connection to UniversalContactServer. {0}", e.ToString()));
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
        public string DownloadAndSaveAttachment(UniversalContactServerProtocol ucsConnection, string documentId, string documentName)
        {
			string path = Path.Combine(OutputFolderName, documentName);

			RequestGetDocument request = new RequestGetDocument();
			request.DocumentId = documentId;
			request.IncludeBinaryContent = true;
			request.DataSource = new NullableDataSourceType(Model.Dst);

			EventGetDocument eventGetDoc = (EventGetDocument)ucsConnection.Request(request);
			if (eventGetDoc == null)
			{
				ShowAndLogErrorMsg(String.Format("Attachment download from the UniversalContactServer failed. Save operation skipped."));
				return null;
			}

			try
			{
				File.WriteAllBytes(path, eventGetDoc.Content);
				if (!Model.EmailPartsInfoStored) Model.EmailPartsPath.Add(path);
			}
			catch (Exception exception)
            {
				ShowAndLogErrorMsg(String.Format("Exception at saving attachment. Path: {0}. {1}\nOperation skipped.", path, exception.Message));
				return null;
            }
            return null;
        }

		/** \brief removes special and whitespace chars from filename 
		 *  \param src source string
		 */
		private string RemoveSpecialChars(string src)
		{
			src = src.Replace("\t", "  ");
			char[] chars = {'<', '>', ':', '\"', '\'', '/','\\', '|', '?', '*'}; // these chars are not allowed in filenames and path names
			char[] invalidPathChars = Path.GetInvalidPathChars();
			var sb = new StringBuilder();
			for (int i = 0; i < src.Length; i++)
			{
		        if (!chars.Contains(src[i]) && !invalidPathChars.Contains(src[i]))
					sb.Append(src[i]);
			}
			string res = sb.ToString();
			res = (res.Length > 0 && (!string.IsNullOrWhiteSpace(res))) ? res : "Empty Subject";
			return res;
		}

		private void ShowAndLogErrorMsg(string s)
		{
			MessageBox.Show(METHOD_NAME + s, "Attention");
			log.Error(METHOD_NAME + s);
		}

		private void ShowAndLogInfoMsg(string s)
		{
			MessageBox.Show(METHOD_NAME + s, "Attention");
			log.Info(METHOD_NAME + s);
		}
	}
}
