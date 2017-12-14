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
    public partial class CSVConnectorWelcomePage : WizardWelcomePageBase
    {
        private CSVConnectorWizardData csvConnectorWizardData = null;

        public CSVConnectorWelcomePage(WizardData wizardData)
        {
            if (wizardData == null)
            {
                throw new ArgumentNullException("wizardData");
            }

            InitializeComponent();

            this.DataContext = wizardData;
            this.csvConnectorWizardData = this.DataContext as CSVConnectorWizardData;

            this.Title = "Before You Begin";
            this.FinishButtonText = "Create";
        }
    }
}