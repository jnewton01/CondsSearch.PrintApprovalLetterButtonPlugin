using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using EllieMae.Encompass.BusinessObjects.Loans;
using EllieMae.Encompass.Automation;
using EllieMae;
using EllieMae.Encompass;
using System.Data;

namespace ConditionsCounters
{
    /// <summary>
    /// Summary description for MonitorWindow.
    /// </summary>
    public class MonitorWindow : System.Windows.Forms.Form

    {
        private DataGridView dataGridView1;
        private BindingSource templateConditionBindingSource;
        private BindingSource conditionsCountersPluginBindingSource;
        private TextBox textBox1;
        private PictureBox pictureBox1;
        private ListBox listBox1;
        private Button addBtn;
        private ComboBox comboBox1;
        private Label label1;
        private Panel panel1;
        private DataGridViewImageColumn dataGridViewImageColumn1;
        private DataGridViewImageColumn Add;
        private Label label2;
        private Label label3;
        private Label label4;
        private IContainer components;

        public MonitorWindow()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MonitorWindow));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Add = new System.Windows.Forms.DataGridViewImageColumn();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.addBtn = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.dataGridViewImageColumn1 = new System.Windows.Forms.DataGridViewImageColumn();
            this.templateConditionBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.conditionsCountersPluginBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.templateConditionBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.conditionsCountersPluginBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.WhiteSmoke;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.Black;
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle4;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.ActiveCaption;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle5;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Add});
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.ControlLightLight;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle6;
            this.dataGridView1.Location = new System.Drawing.Point(12, 60);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(1071, 280);
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick_1);
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // Add
            // 
            this.Add.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Add.HeaderText = "";
            this.Add.Image = ((System.Drawing.Image)(resources.GetObject("Add.Image")));
            this.Add.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Stretch;
            this.Add.Name = "Add";
            this.Add.ReadOnly = true;
            this.Add.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Add.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.Add.ToolTipText = "Insert Condition";
            this.Add.Width = 19;
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(761, 22);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(307, 20);
            this.textBox1.TabIndex = 2;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // listBox1
            // 
            this.listBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBox1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(13, 372);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(1070, 173);
            this.listBox1.TabIndex = 4;
            this.listBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.listBox1_KeyDown);
            // 
            // addBtn
            // 
            this.addBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.addBtn.Location = new System.Drawing.Point(13, 561);
            this.addBtn.Name = "addBtn";
            this.addBtn.Size = new System.Drawing.Size(109, 23);
            this.addBtn.TabIndex = 5;
            this.addBtn.Text = "Add Conditions";
            this.addBtn.UseVisualStyleBackColor = true;
            this.addBtn.Click += new System.EventHandler(this.addBtn_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "",
            "Assets",
            "Affinity",
            "Appraisal",
            "Bond",
            "Closing",
            "Credit",
            "DPA",
            "FHA",
            "GB",
            "Identity",
            "Income",
            "MCC",
            "MSHDA",
            "Property",
            "RCF",
            "RD",
            "ST",
            "Title",
            "VA"});
            this.comboBox1.Location = new System.Drawing.Point(159, 21);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(226, 21);
            this.comboBox1.TabIndex = 6;
            this.comboBox1.SelectionChangeCommitted += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(155, 25);
            this.label1.TabIndex = 7;
            this.label1.Text = "Condition View";
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.Transparent;
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(12, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1071, 54);
            this.panel1.TabIndex = 8;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.BackgroundImage = global::ApprovalsPlugin.Properties.Resources.search_green_dark_icon;
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox1.Location = new System.Drawing.Point(732, 22);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(23, 20);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // dataGridViewImageColumn1
            // 
            this.dataGridViewImageColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.dataGridViewImageColumn1.HeaderText = "";
            this.dataGridViewImageColumn1.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Stretch;
            this.dataGridViewImageColumn1.Name = "dataGridViewImageColumn1";
            this.dataGridViewImageColumn1.ReadOnly = true;
            this.dataGridViewImageColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewImageColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.dataGridViewImageColumn1.ToolTipText = "Insert Condition";
            // 
            // templateConditionBindingSource
            // 
            this.templateConditionBindingSource.DataSource = typeof(EllieMae.Encompass.BusinessObjects.Loans.Logging.TemplateCondition);
            // 
            // conditionsCountersPluginBindingSource
            // 
            this.conditionsCountersPluginBindingSource.DataSource = typeof(ConditionsSearch.ConditionsCountersPlugin);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(13, 356);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(362, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "* To delete a condition select the condition below and press the delete key.";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(161, 5);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(155, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "* Select a filter by option below.";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(868, 5);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(200, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "* Type below to begin filtering conditions.";
            // 
            // MonitorWindow
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.BackColor = System.Drawing.SystemColors.Control;
            this.BackgroundImage = global::ApprovalsPlugin.Properties.Resources.lightbackground1;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1092, 596);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.addBtn);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.dataGridView1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MonitorWindow";
            this.Text = "Gold Star - Add Conditions";
            this.TopMost = true;
            this.Closing += new System.ComponentModel.CancelEventHandler(this.MonitorWindow_Closing);
            this.Load += new System.EventHandler(this.MonitorWindow_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.templateConditionBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.conditionsCountersPluginBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private void MonitorWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Disconnect events
            EncompassApplication.CurrentLoan.FieldChange -= new FieldChangeEventHandler(CurrentLoan_FieldChange);
        }

        private void MonitorWindow_Load(object sender, System.EventArgs e)
        {
            
           
            DataTable tblname = new DataTable();
          
            tblname.Columns.Add("Conditions", typeof(string));
          

            DataRow dr;
                    
            foreach (EllieMae.Encompass.BusinessObjects.Loans.Templates.ConditionTemplate cond in EncompassApplication.Session.Loans.Templates.UnderwritingConditions)
            {

                dr = tblname.NewRow();
               
                dr["Conditions"] = cond.Title.ToString();

                tblname.Rows.Add(dr);
            }

            dataGridView1.DataSource = tblname;
            dataGridView1.AutoResizeColumns();
            dataGridView1.Refresh();
        }

        private void CurrentLoan_LogEntryChange(object source, LogEntryEventArgs e)
        {


        }


        private void CurrentLoan_LogEntryRemoved(object source, LogEntryEventArgs e)
        {

        }

        private void CurrentLoan_LogEntryAdded(object source, LogEntryEventArgs e)
        {



        }


        // Watch for changes to fields
        private void CurrentLoan_FieldChange(object source, FieldChangeEventArgs e)
        {

        }

        private void lvwChanges_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            DataTable tblname = new DataTable();
            tblname.Columns.Add("Conditions", typeof(string));

            DataRow dr;
         

            foreach (EllieMae.Encompass.BusinessObjects.Loans.Templates.ConditionTemplate cond in EncompassApplication.Session.Loans.Templates.UnderwritingConditions)
            {
     
                dr = tblname.NewRow();
              
                dr["Conditions"] = cond.Title.ToString();
             
                tblname.Rows.Add(dr);
            }

            DataView dtv = new DataView(tblname);
            dtv.RowFilter = "Conditions LIKE '%" + textBox1.Text + "%'";

            dataGridView1.DataSource = dtv;
            dataGridView1.AutoResizeColumns();
            dataGridView1.Refresh();

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
          
            listBox1.Items.Add(dataGridView1.Rows[e.RowIndex].Cells["Conditions"].Value.ToString());
   
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            foreach(string lstC in listBox1.Items)
            {
                EllieMae.Encompass.BusinessObjects.Loans.Templates.UnderwritingConditionTemplate conTmpl = EncompassApplication.Session.Loans.Templates.UnderwritingConditions.GetTemplateByTitle(lstC);
                
                EncompassApplication.CurrentLoan.Log.UnderwritingConditions.AddFromTemplate(conTmpl).Source ="GS Search Tool";
              
            }
            this.Close();
        }

        private void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
           if(e.KeyCode == Keys.Delete)
            {
                listBox1.Items.Remove(listBox1.SelectedItem);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable tblname = new DataTable();
            tblname.Columns.Add("Conditions", typeof(string));

            DataRow dr;


            foreach (EllieMae.Encompass.BusinessObjects.Loans.Templates.ConditionTemplate cond in EncompassApplication.Session.Loans.Templates.UnderwritingConditions)
            {

                dr = tblname.NewRow();

                dr["Conditions"] = cond.Title.ToString();

                tblname.Rows.Add(dr);
            }

            DataView dtv = new DataView(tblname);
            dtv.RowFilter = "Conditions LIKE '%" + comboBox1.SelectedItem + "%'";

            dataGridView1.DataSource = dtv;
            dataGridView1.AutoResizeColumns();
            dataGridView1.Refresh();
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
