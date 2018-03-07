﻿using Genesyslab.Platform.Contacts.Protocols;
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
    class FileWriter
    {
        static void Main(string[] args)
        {
            //const int MMF_MAX_SIZE = 1024;  // allocated memory for this memory mapped file (bytes)
            //const int MMF_VIEW_SIZE = 1024; // how many bytes of the allocated memory can this process access

            // creates the memory mapped file
            //MemoryMappedFile mmf = MemoryMappedFile.OpenExisting("mmf1");
            //MemoryMappedViewStream mmvStream = mmf.CreateViewStream(0, MMF_VIEW_SIZE); // stream used to read data

            //BinaryFormatter formatter = new BinaryFormatter();

            //// needed for deserialization
            //byte[] buffer = new byte[MMF_VIEW_SIZE];

            //Message message1;

            //// reads every second what's in the shared memory
            //while (mmvStream.CanRead)
            //{
            //    // stores everything into this buffer
            //    mmvStream.Read(buffer, 0, MMF_VIEW_SIZE);

            //    // deserializes the buffer & prints the message
            //    message1 = (Message)formatter.Deserialize(new MemoryStream(buffer));
            //    Console.WriteLine(message1.title + "\n" + message1.content + "\n");

            //    System.Threading.Thread.Sleep(1000);
            //}

            // connection to UCS contact server
            //ucsConnection = new UniversalContactServerProtocol(new Endpoint("UniversalContactServer", "ling.lauteri.inter", 5130));
            UniversalContactServerProtocol ucsConnection = new UniversalContactServerProtocol(new Endpoint("eS_UniversalContactServer", "genesys1", 6120));
            //ucsConnection = new UniversalContactServerProtocol(new Endpoint(appName, host, port));

            //ucsConnection.Opened += new EventHandler(ucsConnection_Opened);
            //ucsConnection.Error += new EventHandler(ucsConnection_Error);
            //ucsConnection.Closed += new EventHandler(ucsConnection_Closed);

            const int MMF_VIEW_SIZE = 4096;
            MMF_Message message = null;

            System.Diagnostics.Debugger.Break();

            using (MemoryMappedFile mmf = MemoryMappedFile.OpenExisting("adventus_memfile"))
            {
                using (MemoryMappedViewStream stream = mmf.CreateViewStream(0, MMF_VIEW_SIZE))
                {
                    BinaryFormatter formatter = new BinaryFormatter();

                    // needed for deserialization
                    byte[] buffer = new byte[MMF_VIEW_SIZE];

                    stream.Read(buffer, 0, MMF_VIEW_SIZE);

                    // deserializes the buffer & prints the message
                    message = (MMF_Message)formatter.Deserialize(new MemoryStream(buffer));
                    //Console.WriteLine(message.title + "\n" + message.content + "\n");
                }
            }

            RequestGetInteractionContent request = new RequestGetInteractionContent();
            //request.InteractionId = enterpriseEmailInteraction.Id;
            //request.InteractionId = "000D0aD8NBBM001Q";
            if (message != null)
            {
                try
                {
                    ucsConnection.Open();
                }
                catch (Exception e)
                {
                    //ShowAndLogErrorMsg(String.Format("Connection to UniversalContactServer failed. Email saving terminated: {0}", e.ToString()));
                    //ucsConnection.Opened -= new EventHandler(ucsConnection_Opened);
                    //ucsConnection.Error -= new EventHandler(ucsConnection_Error);
                    //ucsConnection.Closed -= new EventHandler(ucsConnection_Closed);
                    //return true;    // stop execution of command chain
                }

                request.InteractionId = message.IntractionId;
                request.IncludeBinaryContent = true;
                request.IncludeAttachments = true;
                //request.DataSource = new NullableDataSourceType(Model.Dst);
                request.DataSource = new NullableDataSourceType(message.DataSourceType == 0 ? Genesyslab.Platform.Contacts.Protocols.ContactServer.DataSourceType.Main : Genesyslab.Platform.Contacts.Protocols.ContactServer.DataSourceType.Archive);

                //GC.Collect(); // to avoid outofmemory exceptions
                EventGetInteractionContent eventGetIxnContent = (EventGetInteractionContent)ucsConnection.Request(request);

                FileWriter pr = new FileWriter();

                if (eventGetIxnContent == null)
                {
                    //ShowAndLogErrorMsg("Request to UniversalContactServer failed. Email saving terminated.");
                    //pr.CloseUCSConnection(ucsConnection);
                    //return true;    // stop execution of the command chain
                }

                AttachmentList attachmentList = eventGetIxnContent.Attachments;
                InteractionContent interactionContent = eventGetIxnContent.InteractionContent;
                pr.CloseUCSConnection(ucsConnection);
                string s = pr.SaveEMLBinaryContent(interactionContent, message.path);
                if (s != String.Empty)
                    Console.WriteLine(s);
            }
            //Environment.Exit(0);
        }

        private void CloseUCSConnection(UniversalContactServerProtocol ucsConnection)
        {
            if (ucsConnection.State != ChannelState.Closed && ucsConnection.State != ChannelState.Closing)
            {
                //ucsConnection.Opened -= new EventHandler(ucsConnection_Opened);
                //ucsConnection.Error -= new EventHandler(ucsConnection_Error);
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
                //if (!Model.EmailPartsInfoStored) Model.EmailPartsPath.Add(path);
            }
            catch (Exception ex)
            {
                //ShowAndLogErrorMsg(String.Format("Cannot Save File {0}: {1}", path, ex.ToString()));
                return ex.Message;
            }
            return String.Empty;
        }



        //private void ShowAndLogErrorMsg(string s)
        //{
        //    MessageBox.Show(METHOD_NAME + s, "Attention");
        //    log.Error(METHOD_NAME + s);
        //}

        //private void ShowAndLogInfoMsg(string s)
        //{
        //    MessageBox.Show(METHOD_NAME + s, "Attention");
        //    log.Info(METHOD_NAME + s);
        //}
    }


    //[Serializable]
    //class MMF_Message
    //{
    //    public string IntractionId { get; set; }
    //    public int DataSourceType { get; set; } // 0 - main, 1 - archive
    //}
}
