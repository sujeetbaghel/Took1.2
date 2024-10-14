using System.Drawing;
using System.Windows.Forms;
using XrmToolBox.Extensibility;

namespace CRMTool
{
    partial class MyPluginControl
    {
        private System.ComponentModel.IContainer components = null;


        /// <param name="disposing">
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyPluginControl));
            this.toolStripMenu = new System.Windows.Forms.ToolStrip();
            this.tsbClose = new System.Windows.Forms.ToolStripButton();
            this.tssSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.tsbDownload = new System.Windows.Forms.ToolStripButton();
            this.btnChooseFile = new System.Windows.Forms.Button();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.btnFetchData = new System.Windows.Forms.Button();  // New button
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.txtEntityLog = new System.Windows.Forms.TextBox();
            this.txtStatus = new System.Windows.Forms.TextBox();
            this.toolStripMenu.SuspendLayout();
            this.SuspendLayout();

            // ToolStripMenu setup
            this.toolStripMenu.Renderer = new HoverToolStripRenderer(); // Apply custom renderer
            this.toolStripMenu.ImageScalingSize = new System.Drawing.Size(28, 28);
            this.toolStripMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
    this.tsbClose,
    this.tssSeparator1,
    this.tsbDownload});
            this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
            this.toolStripMenu.Name = "toolStripMenu";
            this.toolStripMenu.Size = new System.Drawing.Size(1025, 44);
            this.toolStripMenu.TabIndex = 0;

            // tsbClose Button setup
            this.tsbClose.Image = ((System.Drawing.Image)(resources.GetObject("tsbClose.Image")));
            this.tsbClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.tsbClose.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.tsbClose.Margin = new System.Windows.Forms.Padding(11, 0, 4, 0);
            this.tsbClose.Name = "tsbClose";
            this.tsbClose.Padding = new System.Windows.Forms.Padding(5);
            this.tsbClose.Size = new System.Drawing.Size(100, 44);
            this.tsbClose.Text = "Close";
            this.tsbClose.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.tsbClose.Click += new System.EventHandler(this.tsbClose_Click);
            this.tsbClose.Font = new Font("Arial", 9);

            // tssSeparator1
            this.tssSeparator1.Name = "tssSeparator1";
            this.tssSeparator1.Size = new System.Drawing.Size(6, 44);

            // Setup for the new "Fetch Data from Another Environment" button
            this.btnFetchData.BackColor = System.Drawing.Color.FromArgb(34, 177, 76);
            this.btnFetchData.FlatAppearance.BorderSize = 0;
            this.btnFetchData.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFetchData.ForeColor = System.Drawing.Color.White;
            this.btnFetchData.Location = new System.Drawing.Point(440, 62);
            this.btnFetchData.Margin = new System.Windows.Forms.Padding(4);
            this.btnFetchData.Name = "btnSubmit";
            this.btnFetchData.Size = new System.Drawing.Size(140, 30);
            this.btnFetchData.TabIndex = 2;
            this.btnFetchData.Text = "Fetch MetaData";
            this.btnFetchData.UseVisualStyleBackColor = false;
            this.btnFetchData.Font = new Font("Arial", 9);
            this.btnFetchData.Click += new System.EventHandler(this.btnFetchData_Click);

  

            // tsbDownload Button setup
            this.tsbDownload.Image = ((System.Drawing.Image)(resources.GetObject("tsbDownload.Image")));
            this.tsbDownload.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.tsbDownload.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.tsbDownload.Margin = new System.Windows.Forms.Padding(4);
            this.tsbDownload.Name = "tsbDownload";
            this.tsbDownload.Padding = new System.Windows.Forms.Padding(5);
            this.tsbDownload.Size = new System.Drawing.Size(500, 44);
            this.tsbDownload.Text = "Download Sample";
            this.tsbDownload.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.tsbDownload.Click += new System.EventHandler(this.tsbDownload_Click);
            this.tsbDownload.Font = new Font("Arial", 9);

            // btnChooseFile Button setup
            this.btnChooseFile.BackColor = System.Drawing.Color.FromArgb(54, 57, 63);
            this.btnChooseFile.FlatAppearance.BorderSize = 0;
            this.btnChooseFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnChooseFile.ForeColor = System.Drawing.Color.White;
            this.btnChooseFile.Location = new System.Drawing.Point(24, 62);
            this.btnChooseFile.Margin = new System.Windows.Forms.Padding(4);
            this.btnChooseFile.Name = "btnChooseFile";
            this.btnChooseFile.Size = new System.Drawing.Size(140, 30);
            this.btnChooseFile.TabIndex = 1;
            this.btnChooseFile.Text = "Choose Excel File";
            this.btnChooseFile.Font = new Font("Arial", 9);
            this.btnChooseFile.UseVisualStyleBackColor = false;
            this.btnChooseFile.Click += new System.EventHandler(this.btnChooseFile_Click);

            // btnSubmit Button setup
            this.btnSubmit.BackColor = System.Drawing.Color.FromArgb(0, 120, 215);
            this.btnSubmit.FlatAppearance.BorderSize = 0;
            this.btnSubmit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSubmit.ForeColor = System.Drawing.Color.White;
            this.btnSubmit.Location = new System.Drawing.Point(230, 62);
            this.btnSubmit.Margin = new System.Windows.Forms.Padding(4);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(140, 30);
            this.btnSubmit.TabIndex = 2;
            this.btnSubmit.Text = "Create Entity";
            this.btnSubmit.UseVisualStyleBackColor = false;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            this.btnSubmit.Font = new Font("Arial", 9);

            // txtFilePath setup
            this.txtFilePath.BackColor = this.BackColor;
            this.txtFilePath.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtFilePath.Location = new System.Drawing.Point(24, 131);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(4);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.ForeColor = Color.Black; // Optional: Set text color
            this.txtFilePath.Font = new Font("Arial", 10); // Font size and style
            this.txtFilePath.Size = new System.Drawing.Size(800, 29); // Increased width to 800
            this.txtFilePath.TabIndex = 3;
            this.txtFilePath.TextChanged += new System.EventHandler(this.txtFilePath_TextChanged);

            // txtStatus setup
            this.txtStatus.BackColor = this.BackColor;
            this.txtStatus.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtStatus.Location = new System.Drawing.Point(24, 170); // Adjusted location to be just below txtFilePath
            this.txtStatus.Margin = new System.Windows.Forms.Padding(4);
            this.txtStatus.ReadOnly = true;
            this.txtStatus.ForeColor = System.Drawing.Color.Black;
            this.txtStatus.Font = new System.Drawing.Font("Arial", 10);
            this.txtStatus.Size = new System.Drawing.Size(800, 29);
            this.txtStatus.TabIndex = 5;
            this.txtStatus.Text = "";  // Initially empty

            // txtEntityLog setup
            this.txtEntityLog.Font = new System.Drawing.Font("Arial", 9F); // Ensure consistent font size
            this.txtEntityLog.Location = new System.Drawing.Point(24, 210); // Adjusted location to be below txtStatus
            this.txtEntityLog.Margin = new System.Windows.Forms.Padding(4);
            this.txtEntityLog.Multiline = true;
            this.txtEntityLog.Name = "txtEntityLog";
            this.txtEntityLog.ReadOnly = true;
            this.txtEntityLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtEntityLog.Size = new System.Drawing.Size(950, 350);
            this.txtEntityLog.TabIndex = 4;
            this.txtEntityLog.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle; // Add border style
            this.txtEntityLog.BackColor = System.Drawing.Color.White; // Optional: set background color
            this.txtEntityLog.ForeColor = System.Drawing.Color.Black; // Optional: set text color

            // MyPluginControl setup
            this.Controls.Add(this.toolStripMenu);
            this.Controls.Add(this.btnChooseFile);
            this.Controls.Add(this.btnSubmit);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.txtStatus);
            this.Controls.Add(this.txtEntityLog);
            // Add the new button to the form's controls
            this.Controls.Add(this.btnFetchData);
            this.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.Name = "MyPluginControl";
            this.Size = new System.Drawing.Size(1025, 600);
            this.Load += new System.EventHandler(this.MyPluginControl_Load);
            this.toolStripMenu.ResumeLayout(false);
            this.toolStripMenu.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStripMenu;
        private System.Windows.Forms.ToolStripButton tsbClose;
        private System.Windows.Forms.ToolStripButton tsbDownload;
        private System.Windows.Forms.ToolStripSeparator tssSeparator1;

        private System.Windows.Forms.Button btnChooseFile;
        private System.Windows.Forms.Button btnSubmit;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.TextBox txtEntityLog;
        private System.Windows.Forms.TextBox txtStatus;
        private System.Windows.Forms.Button btnFetchData;

    }

    // Custom Renderer class to handle hover effects
    public class HoverToolStripRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
        {
            if (e.Item.Selected || e.Item.Pressed)
            {
                // Draw a light shadow on hover or press
                e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(224, 224, 224)), e.Item.ContentRectangle);
            }
            else
            {
                // Default background rendering
                base.OnRenderButtonBackground(e);
            }
        }
    }
}
