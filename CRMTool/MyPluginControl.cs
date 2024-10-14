using CRMTool.Services;
using Microsoft.Xrm.Sdk;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Web.Services.Description;
using System.Windows.Forms;
using XrmToolBox.Extensibility;

namespace CRMTool
{
    public partial class MyPluginControl : PluginControlBase
    {
        private IOrganizationService _service;
        private string excelFilePath;

        private EntityCreationService _entityCreationService;
        private EntityValidationService _entityValidationService;
        private EntityMetadataService _entityMetadataService;

        public MyPluginControl()
        {
            InitializeComponent();
            _entityValidationService = new EntityValidationService();
        }

        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            _service = Service;
            _entityCreationService = new EntityCreationService(_service, this);
            _entityMetadataService = new EntityMetadataService(_service);
        }

        private async void btnSubmit_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(excelFilePath))
            {
                MessageBox.Show("Please select an Excel file before submitting.");
                return;
            }

            txtStatus.Text = "Processing...";
            txtStatus.ForeColor = Color.Blue;
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            dynamic entities = ExcelDataLoader.LoadExcelEntities(excelFilePath);

            foreach (var entity in entities)
            {
                List<string> validationErrors;

                UpdateLog($"Validating entity: {entity["schema name"]}...");
                await Task.Delay(500);  // Simulate a slight delay for processing

                // Validate the entity
                if (_entityValidationService.ValidateEntityProperties(entity, out validationErrors))
                {
                    UpdateLog($"Creating entity '{entity["schema name"]}'...");
                    await _entityCreationService.CreateCustomEntityAsync(entity);
                }
                else
                {
                    MessageBox.Show(string.Join(Environment.NewLine, validationErrors));
                }
            }

            stopwatch.Stop();

            txtStatus.Text = $"Task finished in {stopwatch.Elapsed.TotalSeconds:F2} seconds.";
            txtStatus.ForeColor = Color.Green;
        }


        //private async void btnSubmit_Click(object sender, EventArgs e)
        //{
        //    //await _entityMetadataService.FetchEntityMetadataAsync("cr367_");
        //    //string targetOrgUrl = "https://orgf1323e5a.crm8.dynamics.com"; // New environment URL
        //    string filePath = @"C:\Users\rajpo\Documents\sample.xlsx";
        //    string entityid = "5bbc37e3-5030-ef11-8e4f-000d3af27a45";
        //    //string targetAccessToken = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1jN2wzSXo5M2c3dXdnTmVFbW13X1dZR1BrbyIsImtpZCI6Ik1jN2wzSXo5M2c3dXdnTmVFbW13X1dZR1BrbyJ9.eyJhdWQiOiJodHRwczovL29yZzhjMjE0NmYzLmNybTguZHluYW1pY3MuY29tIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvMzRiZDhiZWQtMmFjMS00MWFlLTlmMDgtNGUwYTNmMTE3MDZjLyIsImlhdCI6MTcyODE0MjM1MCwibmJmIjoxNzI4MTQyMzUwLCJleHAiOjE3MjgxNDcxNDMsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVFFBeS84WUFBQUFnTjZ6WG1Md3pjdVNGTEsxSFozRWl6NUtXRHIvZkYwMWptSnRWUXplem5BbTJ3dk5GQ3NodkRuR0hhbUp6VWEwIiwiYW1yIjpbInB3ZCJdLCJhcHBpZCI6IjUxZjgxNDg5LTEyZWUtNGE5ZS1hYWFlLWEyNTkxZjQ1OTg3ZCIsImFwcGlkYWNyIjoiMCIsImdpdmVuX25hbWUiOiJTVVJBSiBCSEFOIEtVU0hXQUhBIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMjIzLjE3OC4yMTAuODUiLCJsb2dpbl9oaW50IjoiTy5DaVJpTTJFNE5tTTVPUzB6WldaakxUUTVNRGN0WW1Wak5pMHhOREkzWmpWbE1ERTNaRFlTSkRNMFltUTRZbVZrTFRKaFl6RXROREZoWlMwNVpqQTRMVFJsTUdFelpqRXhOekEyWXhvVE1qSk5RME15TURFNE0wQmpkV05vWkM1cGJpQjUiLCJuYW1lIjoiU1VSQUogQkhBTiBLVVNIV0FIQSIsIm9pZCI6ImIzYTg2Yzk5LTNlZmMtNDkwNy1iZWM2LTE0MjdmNWUwMTdkNiIsInB1aWQiOiIxMDAzMjAwMjQwNjUzMkFFIiwicmgiOiIwLkFWUUE3WXU5Tk1FcXJrR2ZDRTRLUHhGd2JBY0FBQUFBQUFBQXdBQUFBQUFBQUFDaUFDZy4iLCJzY3AiOiJ1c2VyX2ltcGVyc29uYXRpb24iLCJzdWIiOiJjaUlvUWkwVzVQNHR3SkpCdm91cTMtU2wwelFrX2laeGhvRkpJZnhrTjRVIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IkFTIiwidGlkIjoiMzRiZDhiZWQtMmFjMS00MWFlLTlmMDgtNGUwYTNmMTE3MDZjIiwidW5pcXVlX25hbWUiOiIyMk1DQzIwMTgzQGN1Y2hkLmluIiwidXBuIjoiMjJNQ0MyMDE4M0BjdWNoZC5pbiIsInV0aSI6Ii0xREhlVlVfQ0VxdDF4SGxYaDRoQUEiLCJ2ZXIiOiIxLjAiLCJ4bXNfaWRyZWwiOiIxIDI2In0.dOKD0rsw7KAm_Xm_hL0_ZRhApFP9h4Z_-RjT-vPyQdLVrHfD8MYXk4g11UFFTYCgMdS_KN3mmU1940WTkctZpuGwz-6grAzFW2F8L8uank_GiHCpt176C-GwquFHsB04MitJIMLjFqApip8fLtQPYzR-X9pp60h-20-tiud_q1FmNvcZwxF70gaiGzwHbE9aRB_zQURfC-Ov9vgiKi4_tZmZlYCVJbbiLmbiFuwzHxe1kabp88NSbrnXBg-wmw-y-MPlQSoCesfW-mtofaEhrTl3JACdK5o5lmKdq2qVlmXo-jBQJsAk_0HXkKqd0Wft-Es9W9WrxS5Iw-xsyxUgFw"; // Access token for the new environment
        //    //await _entityMetadataService.CreateEntityInNewEnvironmentAsync(filePath, targetOrgUrl, targetAccessToken);
        //    await _entityMetadataService.FetchAndSaveFullEntityMetadataAsync(entityid, filePath);

        //}

        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Select an Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelFilePath = openFileDialog.FileName;
                    txtFilePath.Text = "Path: " + excelFilePath;
                }
            }
        }


        private void tsbDownload_Click(object sender, EventArgs e)
        {
            string projectDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\"));
            string resourceDir = Path.Combine(projectDir, "Resources");
            string sampleFileName = "CRMEntitySample.xlsx"; 
            string sampleFilePath = Path.Combine(resourceDir, sampleFileName);

            // Check if the file exists
            if (!File.Exists(sampleFilePath))
            {
                MessageBox.Show(resourceDir, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Create SaveFileDialog to allow the user to choose where to save the file
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "Save Sample Excel File";
                saveFileDialog.FileName = sampleFileName;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Copy the file to the selected location
                        File.Copy(sampleFilePath, saveFileDialog.FileName, true);
                        MessageBox.Show("Sample file downloaded successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        private void tsbClose_Click(object sender, EventArgs e)
        {
            this.CloseTool();
        }
        private void txtFilePath_TextChanged(object sender, EventArgs e)
        {
            
        }

        // Event handler for the "Fetch Data" button
        private void btnFetchData_Click(object sender, EventArgs e)
        {
            // Open the FetchDataDialog when the button is clicked
            FetchDataDialog fetchDialog = new FetchDataDialog();
            fetchDialog.ShowDialog();  // Show as a dialog
        }
        public void UpdateLog(string message)
        {
            if (txtEntityLog.InvokeRequired) // Check if we need to invoke on the UI thread
            {
                txtEntityLog.Invoke(new Action(() => UpdateLog(message)));
            }
            else
            {
                // Append the message to the log and scroll to the bottom
                txtEntityLog.AppendText($"{DateTime.Now}: {message}\r\n\r\n"); // Use Environment.NewLine for a new line
                txtEntityLog.ScrollToCaret(); // Scroll to the end of the text box
            }
        }




    }
}
