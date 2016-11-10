using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Genesyslab.Desktop.Infrastructure;
using Genesyslab.Desktop.Infrastructure.Commands;
using Genesyslab.Desktop.Modules.Core.Model.Interactions;
using Genesyslab.Desktop.Modules.OpenMedia.Model.Interactions.Email;
using Genesyslab.Desktop.Modules.OpenMedia.Windows.Interactions.MediaView.Email.InteractionInboundEmailView;
using Genesyslab.Desktop.WPFCommon;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;

namespace Adventus.Modules.Email
{
/** \class SaveAttachmentsView
 *  \brief Interaction logic for SaveAttachmentsView.xaml
 */
    public partial class SaveAttachmentsView : UserControl, ISaveAttachmentsView
    {
        readonly IObjectContainer container;
        public object Context { get; set; }

        public SaveAttachmentsView(IObjectContainer container, ISaveAttachmentsViewModel model)
        {
            this.container = container;
            this.Model = model;
            InitializeComponent();
        }
        public ISaveAttachmentsViewModel Model
        {
            get { return this.DataContext as ISaveAttachmentsViewModel; }
            set { this.DataContext = value; }
        }

/** \brief Executed once, at the view object creation
 */
        public void Create()
        {
            IDictionary<string, object> contextDictionary = Context as IDictionary<string, object>;
            Model.Interaction = contextDictionary.TryGetValue("Interaction") as IInteraction;
            IInteractionEmail interactionEmail = Model.Interaction as IInteractionEmail;
            if (interactionEmail == null)
            {
                MessageBox.Show("Interaction is not of IInteractionEmail type");
            }

            switch (Model.Interaction.EntrepriseInteractionCurrent.IdType.Direction)
            {
                case Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.Out:
					Model.SaveButtonVisibility = Visibility.Collapsed;
                    break;
                case Genesyslab.Enterprise.Model.Protocol.MediaDirectionType.In:
					Model.SaveButtonVisibility = Visibility.Visible;
                    break;
            }
		}

/** \brief Executed once, at the view object destruction
 */
        public void Destroy()
        {
        }

/** \brief Event handler
 */
        private void SaveAttachmentsButton_Click(object sender, RoutedEventArgs e)
        {
            IChainOfCommand Command = container.Resolve<ICommandManager>().GetChainOfCommandByName("SaveAttachments");
            IDictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("Model", Model);
            Command.Execute(parameters);
        }
    }
}
