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


        public MyPluginControl()
        {
            InitializeComponent();
            _entityValidationService = new EntityValidationService();
        }

        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            _service = Service;
            _entityCreationService = new EntityCreationService(_service, this);
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
