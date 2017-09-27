using System;
using Genesyslab.Desktop.Infrastructure.DependencyInjection;
using Genesyslab.Desktop.Modules.Core.Model.Agents;
using Genesyslab.Desktop.Modules.Core.SDK.Configurations;
using Genesyslab.Platform.Commons.Collections;
using Genesyslab.Platform.ApplicationBlocks.ConfigurationObjectModel.CfgObjects;
using Genesyslab.Platform.ApplicationBlocks.ConfigurationObjectModel.Queries;
using Genesyslab.Platform.Commons.Logging;


namespace Adventus.Modules.Email
{
	public static class Util
	{
		// read configuration options
		// first read application level configuration option, if not set, read user level config option, if not set use default option
		public static string GetConfigurationOption(string section, string option, IObjectContainer container, string METHOD_NAME) 
		{
			string opt = String.Empty;
			ILogger log = container.Resolve<ILogger>();
			IConfigurationService configurationService = container.Resolve<IConfigurationService>();

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
	}
}
