using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CRMTool
{
    public class FetchDataDialog : Form
    {
        private ComboBox cmbFetchMode;
        private Button btnRetrieve;
        private Label lblFetchMode;
        private TextBox txtInput;
        private string placeholderText;

        public FetchDataDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // Label for "Select Fetch Mode"
            this.lblFetchMode = new Label();
            this.lblFetchMode.AutoSize = true;
            this.lblFetchMode.Location = new System.Drawing.Point(20, 20);
            this.lblFetchMode.Name = "lblFetchMode";
            this.lblFetchMode.Size = new System.Drawing.Size(160, 20);
            this.lblFetchMode.Text = "Select Fetch Mode:";
            this.lblFetchMode.Font = new Font("Arial", 10, FontStyle.Bold);

            // ComboBox for fetch mode options
            this.cmbFetchMode = new ComboBox();
            this.cmbFetchMode.FormattingEnabled = true;
            this.cmbFetchMode.Items.AddRange(new object[] {
                "Entity",
                "Prefix",
                "Solution"
            });
            this.cmbFetchMode.Location = new System.Drawing.Point(200, 20);
            this.cmbFetchMode.Name = "cmbFetchMode";
            this.cmbFetchMode.Size = new System.Drawing.Size(220, 28);
            this.cmbFetchMode.Font = new Font("Arial", 10);
            this.cmbFetchMode.BackColor = Color.White;
            this.cmbFetchMode.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbFetchMode.SelectedIndexChanged += new EventHandler(this.cmbFetchMode_SelectedIndexChanged);

            // TextBox for input (multiline)
            this.txtInput = new TextBox();
            this.txtInput.Location = new System.Drawing.Point(200, 60);
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(220, 60);
            this.txtInput.Multiline = true;
            this.txtInput.Font = new Font("Arial", 10);
            this.txtInput.Padding = new Padding(5);
            this.txtInput.ForeColor = Color.Gray;
            this.txtInput.Text = "Enter Entity, Prefix or Solution";

            // Add event handlers for placeholder behavior
            this.txtInput.Enter += new EventHandler(this.txtInput_Enter);
            this.txtInput.Leave += new EventHandler(this.txtInput_Leave);

            // Retrieve Button
            this.btnRetrieve = new Button();
            this.btnRetrieve.BackColor = Color.FromArgb(0, 120, 215);
            this.btnRetrieve.FlatAppearance.BorderSize = 0;
            this.btnRetrieve.FlatStyle = FlatStyle.Flat;
            this.btnRetrieve.ForeColor = Color.White;
            this.btnRetrieve.Location = new System.Drawing.Point(195, 130);
            this.btnRetrieve.Name = "btnRetrieve";
            this.btnRetrieve.Size = new System.Drawing.Size(140, 30);
            this.btnRetrieve.Text = "Retrieve";
            this.btnRetrieve.Font = new Font("Arial", 9);
            this.btnRetrieve.UseVisualStyleBackColor = false;
            this.btnRetrieve.Click += new EventHandler(this.btnRetrieve_Click);

            // Dialog Form setup
            this.ClientSize = new System.Drawing.Size(460, 180);
            this.Controls.Add(this.lblFetchMode);
            this.Controls.Add(this.cmbFetchMode);
            this.Controls.Add(this.txtInput);
            this.Controls.Add(this.btnRetrieve);
            this.Name = "FetchDataDialog";
            this.Text = "Fetch Data Options";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 240, 240);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }

        private void cmbFetchMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePlaceholder();
        }

        private void txtInput_Enter(object sender, EventArgs e)
        {
            if (txtInput.ForeColor == Color.Gray)
            {
                txtInput.Text = "";
                txtInput.ForeColor = Color.Black;
            }
        }

        private void txtInput_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtInput.Text))
            {
                UpdatePlaceholder();
            }
        }

        private void UpdatePlaceholder()
        {
            switch (cmbFetchMode.SelectedItem?.ToString())
            {
                case "Entity":
                    placeholderText = "Enter Entity Name(s) separated by commas (e.g., entity1, entity2)";
                    break;
                case "Prefix":
                    placeholderText = "Enter Prefix (e.g., dev_)";
                    break;
                case "Solution":
                    placeholderText = "Enter Solution Name";
                    break;
                default:
                    placeholderText = "Enter Entity, Prefix or Solution";
                    break;
            }
            txtInput.Text = placeholderText;
            txtInput.ForeColor = Color.Gray;
        }

        private void btnRetrieve_Click(object sender, EventArgs e)
        {
            string selectedMode = cmbFetchMode.SelectedItem?.ToString();
            string inputText = txtInput.ForeColor == Color.Gray ? "" : txtInput.Text;

            if (string.IsNullOrEmpty(selectedMode))
            {
                MessageBox.Show("Please select a fetch mode.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(inputText))
            {
                MessageBox.Show("Please enter a valid Entity, Prefix, or Solution.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Close this dialog and show the progress dialog
            ProgressDialog progressDialog = new ProgressDialog(selectedMode, inputText);
            progressDialog.ShowDialog(); // Show the progress dialog
            this.Close(); // Close the main form after the progress is done
        }
    }

    public class ProgressDialog : Form
    {
        private ProgressBar progressBar;
        private Label lblProgress;
        private string fetchMode;
        private string inputText;

        public ProgressDialog(string fetchMode, string inputText)
        {
            this.fetchMode = fetchMode;
            this.inputText = inputText;

            InitializeComponent();
            FetchDataAsync(); // Trigger the async fetching process
        }

        private void InitializeComponent()
        {
            // Label for progress description
            this.lblProgress = new Label();
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(20, 20);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(200, 20);
            this.lblProgress.Text = "Fetching data... Please wait";
            this.lblProgress.Font = new Font("Arial", 10, FontStyle.Bold);

            // Progress Bar
            this.progressBar = new ProgressBar();
            this.progressBar.Location = new System.Drawing.Point(20, 60);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(400, 25);
            this.progressBar.Minimum = 0;
            this.progressBar.Maximum = 100;
            this.progressBar.Style = ProgressBarStyle.Continuous;

            // ProgressDialog Form setup
            this.ClientSize = new System.Drawing.Size(460, 120);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.progressBar);
            this.Name = "ProgressDialog";
            this.Text = "Fetching Progress";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }

        private async void FetchDataAsync()
        {
            int totalItemsToFetch = 450; // Example number of items
            progressBar.Value = 0;
            progressBar.Maximum = totalItemsToFetch;

            for (int i = 0; i < totalItemsToFetch; i++)
            {
                await Task.Delay(50); // Simulate network delay or data processing
                progressBar.Value = i + 1;
            }

            MessageBox.Show($"Fetching data for {fetchMode}: {inputText} completed successfully.");
            this.Close(); // Close the progress dialog when done
        }
    }
}
