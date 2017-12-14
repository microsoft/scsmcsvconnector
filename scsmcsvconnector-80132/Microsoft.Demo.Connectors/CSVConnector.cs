/*
 CSVConnector.cs
 
 Description:
 CSVConnector.cs contains the classes for implemeting a CSV connector.
 
 Copyright(c) Microsoft.  All rights reserved.
 This code is licensed under the Microsoft Public License.
 http://www.microsoft.com/opensource/licenses.mspx
 
 THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND,
 EITHER EXPRESSED OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES
 OF FITNESS FOR A PARITCULAR PURPOSE, MERCHANTABILITY, OR
 NON-INFRINGEMENT.
 
 Original Author: Travis Wright (twright@microsoft.com)
 Original Creation Date: Dec 30, 2009
 Original Version: 1.0
*/ 

using System;
using System.Collections.Generic;
using System.Windows;
using System.Resources;             //Has the ResourceManager class in it
using System.Drawing;               //Hsa the Bitmap class in it
using System.Windows.Media.Imaging; //Has the BitmapSource and BitmapSizeOptions classes in it
using System.Globalization;         //Has the CultureInfo class in it
using System.ComponentModel;        //Has the INotifyPropertyChanged class in it
using Microsoft.Win32;              //Has the RegistryKey class in it

//Requires Microsoft.EnterpriseManagement.Core reference
using Microsoft.EnterpriseManagement;
using Microsoft.EnterpriseManagement.Common;
using Microsoft.EnterpriseManagement.Configuration;
using Microsoft.EnterpriseManagement.ConnectorFramework;    //Has IncrementalDiscoveryData in it

//Requires Microsoft.EnterpriseManagement.UI.WpfWizardFramework reference
using Microsoft.EnterpriseManagement.UI.WpfWizardFramework; 

//Requires Microsoft.EnterpriseManagement.UI.SdkDataAccess reference
using Microsoft.EnterpriseManagement.UI.SdkDataAccess;      // Has the ConsoleCommand class in it

//Requires Microsoft.EnterpriseManagement.UI.Foundation reference
using Microsoft.EnterpriseManagement.ConsoleFramework;      //Has the NavigationModelNodeBase and NavigationModelNodeTask in it

namespace Microsoft.Demo.Connectors.CSV
{
    public class CSVConnector : ConsoleCommand
    {
        public CSVConnector()
        {
        }

        public override void ExecuteCommand(IList<NavigationModelNodeBase> nodes, NavigationModelNodeTask task, ICollection<string> parameters)
        {
            if(parameters.Contains("Create"))
            {
                WizardStory wizard = new WizardStory();

                //Set the wizard icon and title bar
                ResourceManager rm = new ResourceManager("Microsoft.Demo.Connectors.Resources", typeof(Resources).Assembly);
                Bitmap bitmap = (Bitmap)rm.GetObject("CSVConnector");
                IntPtr ptr = bitmap.GetHbitmap();
                BitmapSource bitmapsource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(ptr, IntPtr.Zero, Int32Rect.Empty,BitmapSizeOptions.FromEmptyOptions());
                wizard.StoryImage = bitmapsource;
                wizard.WizardWindowTitle = "Create CSV Connector";

                //Create the wizard data
                WizardData data = new CSVConnectorWizardData();
                wizard.WizardData = data;
                
                //Add the wizard pages
                wizard.AddLast(new WizardStep("Welcome", typeof(CSVConnectorWelcomePage),wizard.WizardData));
                wizard.AddLast(new WizardStep("Configuration", typeof(CSVConnectorConfigurationPage), wizard.WizardData));
                wizard.AddLast(new WizardStep("Summary", typeof(CSVConnectorSummaryPage), wizard.WizardData));
                wizard.AddLast(new WizardStep("Results", typeof(CSVConnectorResultPage), wizard.WizardData));

                //Create a wizard window and show it
                WizardWindow wizardwindow = new WizardWindow(wizard);
                // this is needed so that WinForms will pass messages on to the hosted WPF control
                System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(wizardwindow);
                wizardwindow.ShowDialog();

                //Update the view when done with the wizard so that the new connector shows
                if (data.WizardResult == WizardResult.Success)
                {
                    RequestViewRefresh();
                }
            }
            else if (parameters.Contains("Edit"))
            {
                //Get the server name to connect to and connect to the server
                String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
                EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);
                
                //Get the object using the selected node ID
                String strID = String.Empty;
                foreach (NavigationModelNodeBase node in nodes)
                {
                    strID = node["$Id$"].ToString();
                }
                EnterpriseManagementObject emoCSVConnector = emg.EntityObjects.GetObject<EnterpriseManagementObject>(new Guid(strID), ObjectQueryOptions.Default);

                //Create a new "wizard" (also used for property dialogs as in this case), set the title bar, create the data, and add the pages
                WizardStory wizard = new WizardStory();
                wizard.WizardWindowTitle = "Edit CSV Connector";
                WizardData data = new CSVConnectorWizardData(emoCSVConnector);
                wizard.WizardData = data;
                wizard.AddLast(new WizardStep("Configuration", typeof(CSVConnectorConfigurationPage), wizard.WizardData));

                //Show the property page
                PropertySheetDialog wizardWindow = new PropertySheetDialog(wizard);
                
                //Update the view when done so the new values are shown
                bool? dialogResult = wizardWindow.ShowDialog();
                if (dialogResult.HasValue && dialogResult.Value)
                {
                    RequestViewRefresh();
                }
            }
            else if (parameters.Contains("Delete") || parameters.Contains("Disable") || parameters.Contains("Enable"))
            {
                //Get the server name to connect to and create a connection
                String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
                EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);

                //Get the object using the selected node ID
                String strID = String.Empty;
                foreach (NavigationModelNodeBase node in nodes)
                {
                    strID = node["$Id$"].ToString();
                }
                EnterpriseManagementObject emoCSVConnector = emg.EntityObjects.GetObject<EnterpriseManagementObject>(new Guid(strID), ObjectQueryOptions.Default);

                if(parameters.Contains("Delete"))
                {
                    //Remove the object from the database
                    IncrementalDiscoveryData idd = new IncrementalDiscoveryData();
                    idd.Remove(emoCSVConnector);
                    idd.Commit(emg);
                }

                //Get the rule using the connector ID
                ManagementPack mpConnectors = emg.GetManagementPack("Microsoft.Demo.Connectors", null, new Version("1.0.0.0"));
                ManagementPackClass classCSVConnector = mpConnectors.GetClass("Microsoft.Demo.Connectors.CSVConnector");
                String strConnectorID = emoCSVConnector[classCSVConnector, "Id"].ToString();
                ManagementPackRule ruleConnector = mpConnectors.GetRule(strConnectorID);
                
                //Update the Enabled property or delete as appropriate
                if(parameters.Contains("Delete"))
                {
                    ruleConnector.Status = ManagementPackElementStatus.PendingDelete;
                }
                else if (parameters.Contains("Disable"))
                {
                    emoCSVConnector[classCSVConnector, "Enabled"].Value = false;
                    ruleConnector.Enabled = ManagementPackMonitoringLevel.@false;
                    ruleConnector.Status = ManagementPackElementStatus.PendingUpdate;
                }
                else if (parameters.Contains("Enable"))
                {
                    emoCSVConnector[classCSVConnector, "Enabled"].Value = true;
                    ruleConnector.Enabled = ManagementPackMonitoringLevel.@true;
                    ruleConnector.Status = ManagementPackElementStatus.PendingUpdate;
                }

                //Commit the changes to the connector object and rule
                emoCSVConnector.Commit();
                mpConnectors.AcceptChanges();
                
                //Update the view when done so the item is either removed or the updated Enabled value shows
                RequestViewRefresh();
            }
        }
    }

    class CSVConnectorWizardData : WizardData, INotifyPropertyChanged
    {
        #region Variables

        private String strDisplayName = String.Empty;
        private String strDataFilePath = String.Empty;
        private String strMappingFilePath = String.Empty;
        private String strNumberMinutes = String.Empty;
        private String strConnectorID = String.Empty;
        private Guid guidEnterpriseManagementObjectID = Guid.Empty;
        private String strErrorMessage = String.Empty;

        public String DisplayName
        {
            get
            {
                return this.strDisplayName;
            }
            set
            {
                if (this.strDisplayName != value)
                {
                    this.strDisplayName = value;
                    this.NotifyPropertyChanged("DisplayName");
                }
            }
        }

        public String DataFilePath
        {
            get
            {
                return this.strDataFilePath;
            }
            set
            {
                if (this.strDataFilePath != value)
                {
                    this.strDataFilePath = value;
                    this.NotifyPropertyChanged("DataFilePath");
                }
            }
        }

        public String MappingFilePath
        {
            get
            {
                return this.strMappingFilePath;
            }
            set
            {
                if (this.strMappingFilePath != value)
                {
                    this.strMappingFilePath = value;
                    this.NotifyPropertyChanged("MappingFilePath");
                }
            }
        }

        public String NumberMinutes
        {
            get
            {
                return this.strNumberMinutes;
            }
            set
            {
                if (this.strNumberMinutes != value)
                {
                    this.strNumberMinutes = value;
                    this.NotifyPropertyChanged("NumberMinutes");
                }
            }
        }

        public String ConnectorID
        {
            get
            {
                return this.strConnectorID;
            }
            set
            {
                if (this.strConnectorID != value)
                {
                    this.strConnectorID = value;
                    this.NotifyPropertyChanged("RuleName");
                }
            }
        }

        public Guid EnterpriseManagementObjectID
        {
            get
            {
                return this.guidEnterpriseManagementObjectID;
            }
            set
            {
                if (this.guidEnterpriseManagementObjectID != value)
                {
                    this.guidEnterpriseManagementObjectID = value;
                    this.NotifyPropertyChanged("EnterpriseManagementObjectID");
                }
            }
        }

        public String ErrorMessage
        {
            get
            {
                return this.strErrorMessage;
            }
            set
            {
                if (this.strErrorMessage != value)
                {
                    this.strErrorMessage = value;
                    this.NotifyPropertyChanged("ErrorMessage");
                }
            }
        }
        #endregion

        #region Constructors
        internal CSVConnectorWizardData()
        {
        }

        internal CSVConnectorWizardData(EnterpriseManagementObject emoCSVConnector)
        {
            //Get the server name to connect to
            String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();

            //Conneect to the server
            EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);

            ManagementPack mpConnectors = emg.GetManagementPack("Microsoft.Demo.Connectors", null, new Version("1.0.0.0"));
            ManagementPackClass classCSVConnector = mpConnectors.GetClass("Microsoft.Demo.Connectors.CSVConnector");

            this.EnterpriseManagementObjectID = emoCSVConnector.Id;
            this.DisplayName = emoCSVConnector.DisplayName;
            this.DataFilePath = emoCSVConnector[classCSVConnector,"DataFilePath"].ToString();
            this.MappingFilePath = emoCSVConnector[classCSVConnector, "MappingFilePath"].ToString();
            this.NumberMinutes = emoCSVConnector[classCSVConnector, "NumberMinutes"].ToString();
            this.ConnectorID = emoCSVConnector[classCSVConnector, "Id"].ToString();
        }
        #endregion

        public override void AcceptChanges(WizardMode wizardMode)
        {
            if (wizardMode == WizardMode.PropertySheet)
            {
                this.UpdateConnectorInstance();
            }
            else
            {
                try
                {
                    this.CreateConnectorInstance();
                    this.WizardResult = WizardResult.Success;
                }
                catch (Exception ex)
                {
                    this.WizardResult = WizardResult.Failed;
                    this.ErrorMessage = ex.ToString();
                }
            }
        }

        private void CreateConnectorInstance()
        {
            try
            {
                //Get the server name to connect to
                String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();

                //Conneect to the server
                EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);

                //Get the System MP so we can get the system key token and version so we can get other MPs using that info
                ManagementPack mpSystem = emg.ManagementPacks.GetManagementPack(SystemManagementPack.System);
                Version verSystemVersion = mpSystem.Version;
                string strSystemKeyToken = mpSystem.KeyToken;

                //Also get the System Center, Subscription, and Connector Demo MPs - we'll need things from those MPs later
                ManagementPack mpSystemCenter = emg.ManagementPacks.GetManagementPack(SystemManagementPack.SystemCenter);
                ManagementPack mpSubscriptions = emg.GetManagementPack("Microsoft.SystemCenter.Subscriptions", strSystemKeyToken, verSystemVersion);
                ManagementPack mpConnectors = emg.GetManagementPack("Microsoft.Demo.Connectors", null, new Version("1.0.0.0"));

                //Get the CSVConnector class in the Connectors MP
                ManagementPackClass classCSVConnector = mpConnectors.GetClass("Microsoft.Demo.Connectors.CSVConnector");

                //Create a new CreatableEnterpriseManagementObject.  We'll set the properties on this and then post it to the database.
                EnterpriseManagementObject cemoCSVConnector = new CreatableEnterpriseManagementObject(emg, classCSVConnector);

                //Set the property values...

                //Sytem.Entity properties
                cemoCSVConnector[classCSVConnector, "DisplayName"].Value = this.DisplayName;            //Required

                //Microsoft.SystemCenter.Connector properties
                //This is just a tricky way to create a unique ID which conforms to the syntax rules for MP element ID attribute values.
                String strConnectorID = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", "CSVConnector", Guid.NewGuid().ToString("N"));
                cemoCSVConnector[classCSVConnector, "Id"].Value = strConnectorID;                       //Required; Key

                //Unnecessary properties
                //cemoCSVConnector[classCSVConnector, "Description"].Value = "<some description>";      //Optional
                //emoCSVConnector[classCSVConnector, "Name"].Value = "<some name>";                     //Optional           
                //emoCSVConnector[classCSVConnector, "DiscoveryDataIsManaged"].Value = null;            //Optional
                //emoCSVConnector[classCSVConnector, "DiscoveryDataIsShared"].Value = null;             //Optional

                //System.LinkingFramework.DataSource properties            
                cemoCSVConnector[classCSVConnector, "DataProviderDisplayName"].Value = "CSV Connector"; //Optional, shown in Connectors view
                cemoCSVConnector[classCSVConnector, "Enabled"].Value = true;                            //Optional, shown in Connectors view

                //Unncessary properties
                //cemoCSVConnector[classCSVConnector, "DataProviderName"].Value = "CSVConnector";       //Optional
                //cemoCSVConnector[classCSVConnector, "SyncTime"].Value = System.DateTime.Now;          //Optional
                //emoCSVConnector[classCSVConnector,"SolutionName"].Value = "<some string name>"        //Optional
                //emoCSVConnector[classCSVConnector,"ReaderProfileName"].Value = "<SecureRef ID>"       //Optional
                //emoCSVConnector[classCSVConnector,"Reserved"].Value = "<some string>";                //Optional
                //emoCSVConnector[classCSVConnector,"ImpersonationEnabled"].Value = true;               //Optional
                //emoCSVConnector[classCSVConnector,"SyncType"].Value = <enum from SyncTypeEnum>;       //Optional
                //emoCSVConnector[classCSVConnector,"SyncInterval"].Value = 100;                        //Optional
                //emoCSVConnector[classCSVConnector,"SyncNow"].Value = true;                            //Optional

                //Microsoft.Demo.Connectors.CSVConnector properties
                cemoCSVConnector[classCSVConnector, "DataFilePath"].Value = this.DataFilePath;
                cemoCSVConnector[classCSVConnector, "MappingFilePath"].Value = this.MappingFilePath;
                cemoCSVConnector[classCSVConnector, "NumberMinutes"].Value = this.NumberMinutes;

                //Create Connector instance
                cemoCSVConnector.Commit();

                //Now we need to create the CSV Connector rule...

                //Get the Scheduler data source module type from the System MP and the Windows Workflow Task Write Action Module Type from the Subscription MP
                ManagementPackDataSourceModuleType dsmtScheduler = (ManagementPackDataSourceModuleType)mpSystem.GetModuleType("System.Scheduler");
                ManagementPackWriteActionModuleType wamtWindowsWorkflowTaskWriteAction = (ManagementPackWriteActionModuleType)mpSubscriptions.GetModuleType("Microsoft.EnterpriseManagement.SystemCenter.Subscription.WindowsWorkflowTaskWriteAction");

                //Create a new rule for the CSV Connector in the Connectors MP.  Set the name of this rule to be the same as the connector instance ID so there is a pairing between them
                ManagementPackRule ruleCSVConnector = new ManagementPackRule(mpConnectors, strConnectorID);

                //Set the target and other properties of the rule
                ruleCSVConnector.Target = mpSystemCenter.GetClass("Microsoft.SystemCenter.SubscriptionWorkflowTarget");

                //Create a new Data Source Module in the new CSV Connector rule
                ManagementPackDataSourceModule dsmSchedule = new ManagementPackDataSourceModule(ruleCSVConnector, "DS1");

                //Set the configuration of the data source rule.  Pass in the frequency (number of minutes)
                dsmSchedule.Configuration =
                    "<Scheduler>" +
                        "<SimpleReccuringSchedule>" +
                            "<Interval Unit=\"Minutes\">" + this.NumberMinutes + "</Interval>" +
                        "</SimpleReccuringSchedule>" +
                        "<ExcludeDates />" +
                    "</Scheduler>";

                //Set the Schedule Data Source Module Type to the Simple Schedule Module Type from the System MP
                dsmSchedule.TypeID = dsmtScheduler;

                //Add the Scheduler Data Source to the Rule
                ruleCSVConnector.DataSourceCollection.Add(dsmSchedule);

                //Now repeat essentially the same process for the Write Action module...

                //Create a new Write Action Module in the CSV Connector rule
                ManagementPackWriteActionModule wamCSVConnector = new ManagementPackWriteActionModule(ruleCSVConnector, "WA1");

                //Set the Configuration XML
                wamCSVConnector.Configuration =
                    "<Subscription>" +
                        "<WindowsWorkflowConfiguration>" +
                    //Specify the Windows Workflow Foundation workflow Assembly name here
                            "<AssemblyName>CSVConnectorWorkflow</AssemblyName>" +
                    //Specify the type name of the workflow to call in the assembly here:
                            "<WorkflowTypeName>WorkflowAuthoring.CSVConnectorWorkflow</WorkflowTypeName>" +
                            "<WorkflowParameters>" +
                    //Pass in the parameters here.  In this case the two parameters are the data file path and the mapping file path
                                "<WorkflowParameter Name=\"ImportData_DataFilePath\" Type=\"string\">" + this.DataFilePath + "</WorkflowParameter>" +
                                "<WorkflowParameter Name=\"ImportData_FormatFilePath\" Type=\"string\">" + this.MappingFilePath + "</WorkflowParameter>" +
                            "</WorkflowParameters>" +
                            "<RetryExceptions />" +
                            "<RetryDelaySeconds>60</RetryDelaySeconds>" +
                            "<MaximumRunningTimeSeconds>300</MaximumRunningTimeSeconds>" +
                        "</WindowsWorkflowConfiguration>" +
                    "</Subscription>";

                //Set the module type of the module to be the Windows Workflow Task Write Action Module Type from the Subscriptions MP.
                wamCSVConnector.TypeID = wamtWindowsWorkflowTaskWriteAction;

                //Add the Write Action Module to the rule
                ruleCSVConnector.WriteActionCollection.Add(wamCSVConnector);

                //Mark the rule as pending update
                ruleCSVConnector.Status = ManagementPackElementStatus.PendingAdd; ;

                //Accept the rule changes which updates the database
                mpConnectors.AcceptChanges();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + e.InnerException.Message);
            }
        }

        private void UpdateConnectorInstance()
        {
            //Get the server name to connect to and connect
            String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
            EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);

            //Get the Connectors MP and CSV Connector Class
            ManagementPack mpConnectors = emg.GetManagementPack("Microsoft.Demo.Connectors", null, new Version("1.0.0.0"));
            ManagementPackClass classCSVConnector = mpConnectors.GetClass("Microsoft.Demo.Connectors.CSVConnector");

            //Get the Connector object using the object ID
            EnterpriseManagementObject emoCSVConnector = emg.EntityObjects.GetObject<EnterpriseManagementObject>(this.EnterpriseManagementObjectID, ObjectQueryOptions.Default);

            //Set the property values to the new values
            emoCSVConnector[classCSVConnector, "DisplayName"].Value = this.DisplayName;
            emoCSVConnector[classCSVConnector, "DataFilePath"].Value = this.DataFilePath;
            emoCSVConnector[classCSVConnector, "DataFilePath"].Value = this.DataFilePath;
            emoCSVConnector[classCSVConnector, "MappingFilePath"].Value = this.MappingFilePath;
            emoCSVConnector[classCSVConnector, "NumberMinutes"].Value = this.NumberMinutes;

            //Update Connector instance
            emoCSVConnector.Commit();

            //Get the rule using the Connector ID and then update the data source and write action module configuration
            ManagementPackRule ruleConnector = mpConnectors.GetRule(this.ConnectorID);
            ruleConnector.DataSourceCollection[0].Configuration = 
                    "<Scheduler>" +
                        "<SimpleReccuringSchedule>" +
                            "<Interval Unit=\"Minutes\">" + this.NumberMinutes +"</Interval>" + 
                        "</SimpleReccuringSchedule>" + 
                        "<ExcludeDates />" + 
                    "</Scheduler>";

            ruleConnector.WriteActionCollection[0].Configuration =
                    "<Subscription>" +
                        "<WindowsWorkflowConfiguration>" +
                            "<AssemblyName>CSVConnectorWorkflow</AssemblyName>" +
                            "<WorkflowTypeName>WorkflowAuthoring.CSVConnectorWorkflow</WorkflowTypeName>" +
                            "<WorkflowParameters>" +
                                "<WorkflowParameter Name=\"DataFilePath\" Type=\"string\">" + this.DataFilePath + "</WorkflowParameter>" +
                                "<WorkflowParameter Name=\"FormatFilePath\" Type=\"string\">" + this.MappingFilePath + "</WorkflowParameter>" +
                            "</WorkflowParameters>" +
                            "<RetryExceptions />" +
                            "<RetryDelaySeconds>60</RetryDelaySeconds>" +
                            "<MaximumRunningTimeSeconds>300</MaximumRunningTimeSeconds>" +
                        "</WindowsWorkflowConfiguration>" +
                    "</Subscription>";

            ruleConnector.Status = ManagementPackElementStatus.PendingUpdate;
            mpConnectors.AcceptChanges();
        }
        #region PropertyChangedHandler

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged(String info)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(info));
            }
        }
        #endregion
    }
}