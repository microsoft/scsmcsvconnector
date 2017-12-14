 /*
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
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.EnterpriseManagement.UI.WpfWizardFramework;

namespace Microsoft.Demo.Connectors.CSV
{
    /// <summary>
    /// Interaction logic for CSVConnectorConfigurationPage.xaml
    /// </summary>
    public partial class CSVConnectorConfigurationPage : WizardRegularPageBase
    {
        private CSVConnectorWizardData csvConnectorWizardData = null;

        public CSVConnectorConfigurationPage(WizardData wizardData)
        {
            if (wizardData == null)
            {
                throw new ArgumentNullException("wizardData");
            }

            InitializeComponent();

            this.DataContext = wizardData;
            this.csvConnectorWizardData = this.DataContext as CSVConnectorWizardData;

            this.Title = "Configure the CSV Connector";
            this.FinishButtonText = "Create";
        }

        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }
    }
}
