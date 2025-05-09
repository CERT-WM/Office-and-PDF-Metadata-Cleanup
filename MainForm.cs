using System;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;

namespace MetadataCleanerApp
{
    public partial class MainForm : Form
    {
        private string? outputFolder;
        private Button openButton;
        private Button selectFolderButton;
        private Button cleanButton;
        private ListBox fileListBox;
        private Label outputFolderLabel;

        public MainForm()
        {
            // Initialize fields
            openButton = new Button();
            selectFolderButton = new Button();
            cleanButton = new Button();
            fileListBox = new ListBox();
            outputFolderLabel = new Label();

            InitializeComponent();

            // Set the form's icon
            try
            {
                this.Icon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("MetadataCleanerApp.AppIcon.ico"));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load icon: {ex.Message}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void InitializeComponent()
        {
            // Open Button
            this.openButton.Text = "Open Files";
            this.openButton.Location = new System.Drawing.Point(20, 20);
            this.openButton.Size = new System.Drawing.Size(100, 30);
#pragma warning disable CS8622 // Suppress nullability warning for event handler
            this.openButton.Click += new EventHandler(this.OpenButton_Click);
#pragma warning restore CS8622

            // Select Folder Button
            this.selectFolderButton.Text = "Select Output Folder";
            this.selectFolderButton.Location = new System.Drawing.Point(130, 20);
            this.selectFolderButton.Size = new System.Drawing.Size(150, 30);
#pragma warning disable CS8622
            this.selectFolderButton.Click += new EventHandler(this.SelectFolderButton_Click);
#pragma warning restore CS8622

            // Clean Button
            this.cleanButton.Text = "Clean Metadata";
            this.cleanButton.Location = new System.Drawing.Point(300, 20);
            this.cleanButton.Size = new System.Drawing.Size(120, 30);
#pragma warning disable CS8622
            this.cleanButton.Click += new EventHandler(this.CleanButton_Click);
#pragma warning restore CS8622

            // File ListBox
            this.fileListBox.Location = new System.Drawing.Point(20, 60);
            this.fileListBox.Size = new System.Drawing.Size(400, 200);
            this.fileListBox.AllowDrop = true;
#pragma warning disable CS8622
            this.fileListBox.DragEnter += new DragEventHandler(this.FileListBox_DragEnter);
            this.fileListBox.DragDrop += new DragEventHandler(this.FileListBox_DragDrop);
#pragma warning restore CS8622

            // Output Folder Label
            this.outputFolderLabel.Text = "Output Folder: Not selected";
            this.outputFolderLabel.Location = new System.Drawing.Point(20, 270);
            this.outputFolderLabel.Size = new System.Drawing.Size(400, 20);

            // Form Settings
            this.Text = "Metadata Cleaner";
            this.Size = new System.Drawing.Size(450, 350);
            this.Controls.Add(this.openButton);
            this.Controls.Add(this.selectFolderButton);
            this.Controls.Add(this.cleanButton);
            this.Controls.Add(this.fileListBox);
            this.Controls.Add(this.outputFolderLabel);
        }

        private void OpenButton_Click(object? sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Multiselect = true;
                openFileDialog.Filter = "Office and PDF Files|*.docx;*.xlsx;*.pptx;*.pdf|All Files|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (string file in openFileDialog.FileNames)
                    {
                        if (!fileListBox.Items.Contains(file))
                        {
                            fileListBox.Items.Add(file);
                        }
                    }
                }
            }
        }

        private void SelectFolderButton_Click(object? sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select the folder where cleaned files will be saved.";
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    outputFolder = folderBrowserDialog.SelectedPath;
                    outputFolderLabel.Text = $"Output Folder: {outputFolder}";
                }
            }
        }

        private void FileListBox_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void FileListBox_DragDrop(object? sender, DragEventArgs e)
        {
            string[]? files = e.Data?.GetData(DataFormats.FileDrop) as string[];
            if (files != null)
            {
                foreach (string file in files)
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (ext == ".docx" || ext == ".xlsx" || ext == ".pptx" || ext == ".pdf")
                    {
                        if (!fileListBox.Items.Contains(file))
                        {
                            fileListBox.Items.Add(file);
                        }
                    }
                }
            }
        }

        private void CleanButton_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(outputFolder))
            {
                MessageBox.Show("No output folder selected. Please select an output folder.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
                {
                    folderBrowserDialog.Description = "Select the folder where cleaned files will be saved.";
                    if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                    {
                        outputFolder = folderBrowserDialog.SelectedPath;
                        outputFolderLabel.Text = $"Output Folder: {outputFolder}";
                    }
                    else
                    {
                        return; // User canceled folder selection
                    }
                }
            }

            if (fileListBox.Items.Count == 0)
            {
                MessageBox.Show("No files selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            foreach (string file in fileListBox.Items)
            {
                try
                {
                    MetadataCleaner.CleanMetadata(file, outputFolder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error processing {file}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            MessageBox.Show("Metadata cleaning completed.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            fileListBox.Items.Clear();
        }
    }
}