using Genesyslab.Platform.Contacts.Protocols;
using Genesyslab.Platform.Commons.Protocols;
using System;
using Genesyslab.Platform.Contacts.Protocols.ContactServer.Requests;
using Genesyslab.Platform.Contacts.Protocols.ContactServer;
using Genesyslab.Platform.Contacts.Protocols.ContactServer.Events;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Runtime.Serialization.Formatters.Binary;
using Adventus.Modules.Email;

namespace InteractionWorkspaceHelper
{
    public class FileWriter
    {
        public const int MMF_VIEW_SIZE = 2048;
        public static void Main(string[] args)
        {
            
            MMF_Message message = null;
            FileWriter pr = new FileWriter();
            try
            {
                using (MemoryMappedFile mmf = MemoryMappedFile.OpenExisting("adventus_wde_memfile"))
                {
                    using (MemoryMappedViewStream stream = mmf.CreateViewStream(0, MMF_VIEW_SIZE))
                    {
                        BinaryFormatter formatter = new BinaryFormatter();

                        // needed for deserialization
                        byte[] buffer = new byte[MMF_VIEW_SIZE];

                        stream.Read(buffer, 0, MMF_VIEW_SIZE);

                        // deserializes the buffer & prints the message
                        message = (MMF_Message)formatter.Deserialize(new MemoryStream(buffer));
                        if(message == null)
                        {
                            Console.WriteLine("6");
                            return;
                        }
                    }
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("4");
                return;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return;
            }

            UniversalContactServerProtocol ucsConnection = new UniversalContactServerProtocol(new Endpoint(message.UCSappName, message.UCSHost, message.UCSPort));
            //ucsConnectiown.Opened += new EventHandler(pr.ucsConnection_Opened);
            ucsConnection.Error += new EventHandler(pr.ucsConnection_Error);
            //ucsConnection.Closed += new EventHandler(pr.ucsConnection_Closed);

            RequestGetInteractionContent request = new RequestGetInteractionContent();

            try
            {
                ucsConnection.Open();
            }
            catch (Exception)
            {
                Console.WriteLine("1");
                return;
            }

            request.InteractionId = message.IntractionId;
            request.IncludeBinaryContent = true;
            request.IncludeAttachments = false; // attachments have been already saved earlier
            request.DataSource = new NullableDataSourceType(message.DataSourceType == 0 ? Genesyslab.Platform.Contacts.Protocols.ContactServer.DataSourceType.Main : Genesyslab.Platform.Contacts.Protocols.ContactServer.DataSourceType.Archive);

            EventGetInteractionContent eventGetIxnContent = (EventGetInteractionContent)ucsConnection.Request(request);

            if (eventGetIxnContent == null)
            {
                pr.CloseUCSConnection(ucsConnection);
                Console.WriteLine("2");
                return;
            }

            InteractionContent interactionContent = eventGetIxnContent.InteractionContent;
            pr.CloseUCSConnection(ucsConnection);
            string s = pr.SaveEMLBinaryContent(interactionContent, message.path);
            if (s != String.Empty)
            {
                Console.WriteLine("3");
                return;
            }

            Console.WriteLine("0");
            return;
        }

        private void CloseUCSConnection(UniversalContactServerProtocol ucsConnection)
        {
            if (ucsConnection.State != ChannelState.Closed && ucsConnection.State != ChannelState.Closing)
            {
                //ucsConnection.Opened -= new EventHandler(ucsConnection_Opened);
                ucsConnection.Error -= new EventHandler(ucsConnection_Error);
                //ucsConnection.Closed -= new EventHandler(ucsConnection_Closed);
                ucsConnection.Close();
                ucsConnection.Dispose();
            }
        }

        private string SaveEMLBinaryContent(InteractionContent interactionContent, string path)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllBytes(path, interactionContent.Content);
            }
            catch (Exception ex)
            {
                //ShowAndLogErrorMsg(String.Format("Cannot Save File {0}: {1}", path, ex.ToString()));
                return ex.Message;
            }
            return String.Empty;
        }
        //private void ucsConnection_Closed(object sender, EventArgs e)
        //{

        //}

        private void ucsConnection_Error(object sender, EventArgs e)
        {
            IMessage message = ((MessageEventArgs)e).Message;
            Console.WriteLine("5");
            Environment.Exit(0);
        }

        //private void ucsConnection_Opened(object sender, EventArgs e)
        //{

        //}
    }
}
