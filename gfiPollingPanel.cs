using System;
using System.IO;
using System.Net;
using System.Xml;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using CSVFile;
using System.Data.SqlClient;
using System.Globalization;
using System.Collections;


namespace ChoiceGFIBilling
{
    public partial class gfiPollingPanel : UserControl
    {
        const int tpServerMonitoringAsset = 2;
        const int tpPCMonitoringAsset = 3;
        const int priceLevel = 1;
        const int useFlatPrice = 1;
        frmChoiceGFIBilling panelParent;
        SqlDataAdapter tpBillingAdapter;
        String gfiAPIString = "http://www.systemmonitor.us/api/?apikey=###############################&service=";
        String deviceURL = "http://www.systemmonitor.us/dashboard/";
        DataTable gfiClientDataTable = new DataTable("GFI Client IDs");
        DataTable gfiAssets = new DataTable("GFI Assets");
        DateTime currentPollTime;

        public gfiPollingPanel()
        {
            InitializeComponent();
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
            backgroundWorker2.DoWork += new DoWorkEventHandler(backgroundWorker2_DoWork);
            backgroundWorker2.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker2_ProgressChanged);
            backgroundWorker2.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker2_RunWorkerCompleted);

            // Register load event to assign parent once panel is created.
            this.Load += new EventHandler(gfiPollingPanel_OnLoad);

            // Init UI elements
            panelParent = (frmChoiceGFIBilling)ParentForm;

            gfiClientDataTable.Columns.Add("Client Name");
            gfiClientDataTable.Columns.Add("ClientID");

            gfiAssets.Columns.Add("Site ID");
            gfiAssets.Columns.Add("Client Name");
            gfiAssets.Columns.Add("PC Name");
            gfiAssets.Columns.Add("OS");
        }

        public void gfiPollingPanel_OnLoad(object sender, EventArgs e)
        {
            panelParent = (frmChoiceGFIBilling)ParentForm;
        }

        private void flagRemovedAssetsInTigerpaw(DateTime pollTime)
        {
            // Create SQLCommandBuilder that will generate update, insert and delete commands based on SQLDataAdapter.
            SqlCommandBuilder tpAssetCmdBuilder = new SqlCommandBuilder(panelParent.tpAssetAdapter);

            // Create DataSet for TP Assets to compare with GFI Asset IDs
            DataSet tpAssetDataSet = new DataSet();
            panelParent.tpAssetAdapter.Fill(tpAssetDataSet);
            foreach (DataRow assetRow in tpAssetDataSet.Tables[0].Rows)
            {
                // If an AssetID was not found in the poll, it isn't there anymore and needs to be updated as inactive.
                if ((DateTime)assetRow["LastPolledDate"] < pollTime)
                {
                    assetRow["InActiveIndicator"] = 1;
                    assetRow["DateRemoved"] = DateTime.Now;
                    panelParent.tpAssetAdapter.Update(tpAssetDataSet);
                }
            }
        }


        private void deleteAssetsFromTigerpaw()
        {   
            // Create SQLCommandBuilder that will generate update, insert and delete commands based on SQLDataAdapter.
            SqlCommandBuilder tpAssetCmdBuilder = new SqlCommandBuilder(panelParent.tpAssetAdapter);

            // Create DataSet to hold TP Asset table and make changes to it.
            DataSet tpAssetDataSet = new DataSet();
            panelParent.tpAssetAdapter.Fill(tpAssetDataSet);
            foreach (DataRow assetRow in tpAssetDataSet.Tables[0].Rows)
            {
                assetRow.Delete();
            }

            panelParent.tpAssetAdapter.Update(tpAssetDataSet);
        }

        private void updateAssetInTigerpaw(String externalAccount, int assetType, String tpAccount, String assetName, String assetID, String assetOS)
        {
            // Calculate Flat Price based on client contract (tblMSPAgreementAssetTypes)
            String flatPrice = "0";
            DataSet tpCoveredAssetDataSet = new DataSet();
            panelParent.tpCoveredAssetAdapter.Fill(tpCoveredAssetDataSet);
            DataRow[] tpCoveredAssetRows = tpCoveredAssetDataSet.Tables[0].Select("AccountNumber = '" + tpAccount + "' AND FKMSPAssetTypes = '" + assetType + "'");
            if (tpCoveredAssetRows.Count() != 0)
            {
                DataRow tpCoveredAssetRow = tpCoveredAssetRows[0];
                flatPrice = tpCoveredAssetRow["DefaultFlatPrice"].ToString();
            }

            // Create SQLCommandBuilder that will generate update, insert and delete commands based on SQLDataAdapter.
            SqlCommandBuilder tpAssetCmdBuilder = new SqlCommandBuilder(panelParent.tpAssetAdapter);

            // Create DataSet to hold TP Asset table and make changes to it.
            DataSet tpAssetDataSet = new DataSet();
            panelParent.tpAssetAdapter.Fill(tpAssetDataSet);

            // Determine if the asset already exists in Tigerpaw
            DataRow[] tpAssetRows = tpAssetDataSet.Tables[0].Select("ProvidersAssetID LIKE '" + assetID + "'");
            if (tpAssetRows.Count() != 0)
            {
                // Asset Exists, update asset name (PC names can be changed), poll time and pricing
                DataRow tpAssetRow = tpAssetRows[0];
                tpAssetRow["ProvidersAssetName"] = assetName;
                tpAssetRow["DeviceURL"] = deviceURL;
                tpAssetRow["UseFlatPrice"] = useFlatPrice;
                tpAssetRow["LastPolledDate"] = DateTime.Now;
                tpAssetRow["InActiveIndicator"] = 0;
                tpAssetRow["DateRemoved"] = DBNull.Value;
            }
            else
            {
                // New asset, generate TP Asset Row to be inserted into database
                DataRow newSQLRow = tpAssetDataSet.Tables[0].NewRow();
                newSQLRow["FKAssignedExternalAccount"] = externalAccount; // Calculated per client based on the tblAssignedExternalAccounts table
                newSQLRow["FKMSPAssetTypes"] = assetType; // 2 is Monitored Server, 3 is Monitored PC, 4 is Spam Filtering
                newSQLRow["AccountNumber"] = tpAccount;
                newSQLRow["ProvidersAssetName"] = assetName;
                newSQLRow["ProvidersAssetID"] = assetID;
                newSQLRow["DeviceURL"] = deviceURL;
                newSQLRow["PriceLevel"] = priceLevel;
                newSQLRow["FlatPrice"] = flatPrice; // Calculated based on the contract.
                newSQLRow["UseFlatPrice"] = useFlatPrice;
                newSQLRow["PriceOverride"] = 0;
                newSQLRow["DateAdded"] = DateTime.Now;
                newSQLRow["LastPolledDate"] = DateTime.Now;
                newSQLRow["InActiveIndicator"] = 0;
                newSQLRow["DateRemoved"] = DBNull.Value;

                tpAssetDataSet.Tables[0].Rows.Add(newSQLRow);
            }

            panelParent.tpAssetAdapter.Update(tpAssetDataSet);
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            WebClient client = new WebClient();
            DateTime currentDate = DateTime.Now;
            ArrayList clientIDs = new ArrayList();

            // Lookup GFI external account IDs in TP Database
            DataSet tpDataSet = new DataSet("TP Client IDs");
            tpBillingAdapter = new SqlDataAdapter("SELECT tblAccounts.AccountNumber, tblAccounts.AccountName, tblAssignedExternalAccounts.ExternalID, tblAssignedExternalAccounts.Comment FROM tblAccounts, tblAssignedExternalAccounts WHERE tblAccounts.AccountNumber = tblAssignedExternalAccounts.FKAccountNumber", panelParent.tigerpawDBConn);
            tpBillingAdapter.Fill(tpDataSet);

            // Get XML client list from GFI
            String gfiAPIData = client.DownloadString(gfiAPIString + "list_clients");
            XmlReader reader = XmlReader.Create(new StringReader(gfiAPIData));

            // Create array of clients and data view of clients from GFI that do not have TP external connectors
            while (reader.Read())
            {
                if (reader.ReadToFollowing("client"))
                {
                    reader.ReadToFollowing("name");
                    String gfiClientName = reader.ReadElementContentAsString();
                    String gfiClientID = reader.ReadElementContentAsString();

                    // Looks for the GFI ID in the list of accounts with external IDs and adds it to the displayed list if it isn't there
                    DataRow[] tpRows = tpDataSet.Tables[0].Select("ExternalID LIKE '" + gfiClientID + "' AND Comment LIKE 'GFI Max'");
                    if (tpRows.Count() == 0)
                    {
                        gfiClientDataTable.Rows.Add(gfiClientName, gfiClientID);//, tpClientName, tpClientID);
                    }
                    else
                    {
                        // Only add accounts linked to Tigerpaw to the polling queue
                        // Need to determine what the external key is based on the tblAssignedExternalAccounts table
                        DataSet tpExternalAccountsDataSet = new DataSet("TP External Keys");
                        panelParent.tpExternalAccountsAdapter.Fill(tpExternalAccountsDataSet);
                        DataRow tpRow = tpRows[0];
                        String tpAccountNumber = tpRow["AccountNumber"].ToString();
                        String tpAccountName = tpRow["AccountName"].ToString();

                        DataRow[] keyRows = tpExternalAccountsDataSet.Tables[0].Select("ExternalID LIKE '" + gfiClientID + "'");
                        if (keyRows.Count() != 0)
                        {
                            DataRow keyRow = keyRows[0];
                            String tpExternalKey = keyRow["AssignedExternalAccountsKeyId"].ToString();
                            ArrayList clientArray = new ArrayList();
                            clientArray.Add(gfiClientID);
                            clientArray.Add(tpExternalKey);
                            clientArray.Add(tpAccountNumber);
                            clientArray.Add(tpAccountName);
                            clientIDs.Add(clientArray);
                        }
                    }
                }
            }

            if (gfiClientDataTable.Rows.Count == 0)
            {
                gfiClientDataTable.Rows.Add("All GFI clients have links to Tigerpaw", "");
            }
            e.Result = clientIDs;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            panelParent.updateProgressBar(e.ProgressPercentage);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ArrayList clientIDs = new ArrayList();
            clientIDs = (ArrayList)e.Result;
            
            panelParent.updateStatusBarText("Client list update complete.");
            panelParent.changeProgressBarMax(clientIDs.Count);
            dataGridView1.DataSource = gfiClientDataTable;
            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            DataGridViewColumn lastColumn = dataGridView1.Columns.GetLastColumn(DataGridViewElementStates.Visible, DataGridViewElementStates.None);
            lastColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Refresh();
            panelParent.updateStatusBarText("Polling assets from GFI...");
            currentPollTime = DateTime.Now;
            backgroundWorker2.RunWorkerAsync(clientIDs);
        }

		// Background process to poll assets so that the application does not lock up during the long process.
        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            WebClient client = new WebClient();
            DateTime currentDate = DateTime.Now;
            int currentProgress = 0;

            // ASSET RETRIEVAL

            // Create Asset Table for Testing Display
            DataTable gfiAssetDataTable = new DataTable("GFI Assets");
            gfiAssetDataTable.Columns.Add("Client ID");
            gfiAssetDataTable.Columns.Add("Asset Name");
            gfiAssetDataTable.Columns.Add("Asset Type");

            // Getting SiteIDs for each client ID.
            ArrayList clientIDs = (ArrayList)e.Argument;
            String gfiAPIData = "";
            XmlReader reader;
            ArrayList siteIDs = new ArrayList();
            String assetName = "";
            String assetID = "";
            String assetOS = "";

            foreach (ArrayList clientIDArrayList in clientIDs)
            {

                gfiAPIData = client.DownloadString(gfiAPIString + "list_sites&clientid=" + clientIDArrayList[0]);
                reader = XmlReader.Create(new StringReader(gfiAPIData));

                // Create array of site IDs for client.
                siteIDs.Clear();

                while (reader.Read())
                {
                    if (reader.ReadToFollowing("site"))
                    {
                        reader.ReadToFollowing("siteid");
                        siteIDs.Add(reader.ReadElementContentAsString());
                    }
                }

                // Getting server level data.
                for (int i = 0; i < siteIDs.Count; i++)
                {
                    gfiAPIData = client.DownloadString(gfiAPIString + "list_servers&siteid=" + (String)siteIDs[i]);

                    reader = XmlReader.Create(new StringReader(gfiAPIData));

                    while (reader.Read())
                    {
                        if (reader.ReadToFollowing("server"))
                        {
                            reader.ReadToFollowing("name");
                            assetName = reader.ReadElementContentAsString();
                            reader.ReadToFollowing("os");
                            assetOS = reader.ReadElementContentAsString();
                            reader.ReadToFollowing("assetid");
                            assetID = reader.ReadElementContentAsString();

                            updateAssetInTigerpaw((String)clientIDArrayList[1], tpServerMonitoringAsset, (String)clientIDArrayList[2], assetName, assetID, assetOS);
                        }
                    }
                }

                // Getting workstation level data
                for (int i = 0; i < siteIDs.Count; i++)
                {
                    gfiAPIData = client.DownloadString(gfiAPIString + "list_workstations&siteid=" + (String)siteIDs[i]);

                    reader = XmlReader.Create(new StringReader(gfiAPIData));

                    while (reader.Read())
                    {

                        if (reader.ReadToFollowing("workstation"))
                        {
                            reader.ReadToFollowing("name");
                            assetName = reader.ReadElementContentAsString();
                            reader.ReadToFollowing("os");
                            assetOS = reader.ReadElementContentAsString();
                            reader.ReadToFollowing("assetid");
                            assetID = reader.ReadElementContentAsString();


                            gfiAssets.Rows.Add((String)siteIDs[0], clientIDArrayList[3], assetName, assetOS); // Creating a data set to be exported for all PCs.

                            updateAssetInTigerpaw((String)clientIDArrayList[1], tpPCMonitoringAsset, (String)clientIDArrayList[2], assetName, assetID, assetOS);
                        }
                    }
                }
				// Update status bar.
                currentProgress++;
                backgroundWorker2.ReportProgress(currentProgress);
            }
        }

		// Updates status bar with progress bar information.
        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            panelParent.updateProgressBar(e.ProgressPercentage);
        }

		// Updates table with polled assets and updates status bar with completion notification.
        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            flagRemovedAssetsInTigerpaw(currentPollTime);
            dataGridView2.DataSource = gfiAssets;
            dataGridView2.AutoResizeColumns();
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            DataGridViewColumn lastColumn = dataGridView2.Columns.GetLastColumn(DataGridViewElementStates.Visible, DataGridViewElementStates.None);
            lastColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView2.Refresh();
            panelParent.updateStatusBarText("Polling complete.");
            button1.Enabled = true;
        }


		// Polling button updates assets from GFI.
        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            
            int totalRows = gfiClientDataTable.Rows.Count;

            for (int i = 0; i < totalRows; i++)
            {
                gfiClientDataTable.Rows[0].Delete();
            }
            panelParent.updateStatusBarText("Retrieving updated client list from GFI...");
            backgroundWorker1.RunWorkerAsync();
        }


    }
}
