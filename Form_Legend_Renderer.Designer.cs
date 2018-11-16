namespace GSC_Legend_Renderer
{
    partial class Form_Legend_Renderer
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_Legend_Renderer));
            this.label_selectTable = new System.Windows.Forms.Label();
            this.button_selectTable = new System.Windows.Forms.Button();
            this.button_Start = new System.Windows.Forms.Button();
            this.button_Cancel = new System.Windows.Forms.Button();
            this.groupBox_FieldList = new System.Windows.Forms.GroupBox();
            this.comboBox_DescriptionField = new System.Windows.Forms.ComboBox();
            this.label_DescriptionField = new System.Windows.Forms.Label();
            this.comboBox_HeadingField = new System.Windows.Forms.ComboBox();
            this.label_HeadingField = new System.Windows.Forms.Label();
            this.comboBox_Label2StyleField = new System.Windows.Forms.ComboBox();
            this.label_Label2StyleField = new System.Windows.Forms.Label();
            this.comboBox_Label2Field = new System.Windows.Forms.ComboBox();
            this.label_Label2Field = new System.Windows.Forms.Label();
            this.comboBox_Label1StyleField = new System.Windows.Forms.ComboBox();
            this.label_Label1StyleField = new System.Windows.Forms.Label();
            this.comboBox_Label1Field = new System.Windows.Forms.ComboBox();
            this.label_Label1Field = new System.Windows.Forms.Label();
            this.comboBox_Style2Field = new System.Windows.Forms.ComboBox();
            this.label_Style2Field = new System.Windows.Forms.Label();
            this.comboBox_Style1Field = new System.Windows.Forms.ComboBox();
            this.label_Style1Field = new System.Windows.Forms.Label();
            this.comboBox_ElementField = new System.Windows.Forms.ComboBox();
            this.label_ElementField = new System.Windows.Forms.Label();
            this.comboBox_ColumnField = new System.Windows.Forms.ComboBox();
            this.label_ColumnField = new System.Windows.Forms.Label();
            this.label_orderField = new System.Windows.Forms.Label();
            this.comboBox_orderField = new System.Windows.Forms.ComboBox();
            this.checkBox_DEMBoxes = new System.Windows.Forms.CheckBox();
            this.comboBox_SelectTable = new System.Windows.Forms.ComboBox();
            this.checkBox_autoCalculateColumns = new System.Windows.Forms.CheckBox();
            this.groupBox_FieldList.SuspendLayout();
            this.SuspendLayout();
            // 
            // label_selectTable
            // 
            resources.ApplyResources(this.label_selectTable, "label_selectTable");
            this.label_selectTable.Name = "label_selectTable";
            // 
            // button_selectTable
            // 
            resources.ApplyResources(this.button_selectTable, "button_selectTable");
            this.button_selectTable.Name = "button_selectTable";
            this.button_selectTable.UseVisualStyleBackColor = true;
            this.button_selectTable.Click += new System.EventHandler(this.button_selectTable_Click);
            // 
            // button_Start
            // 
            resources.ApplyResources(this.button_Start, "button_Start");
            this.button_Start.Name = "button_Start";
            this.button_Start.UseVisualStyleBackColor = true;
            this.button_Start.Click += new System.EventHandler(this.button_Start_Click);
            // 
            // button_Cancel
            // 
            resources.ApplyResources(this.button_Cancel, "button_Cancel");
            this.button_Cancel.Name = "button_Cancel";
            this.button_Cancel.UseVisualStyleBackColor = true;
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);
            // 
            // groupBox_FieldList
            // 
            resources.ApplyResources(this.groupBox_FieldList, "groupBox_FieldList");
            this.groupBox_FieldList.Controls.Add(this.comboBox_DescriptionField);
            this.groupBox_FieldList.Controls.Add(this.label_DescriptionField);
            this.groupBox_FieldList.Controls.Add(this.comboBox_HeadingField);
            this.groupBox_FieldList.Controls.Add(this.label_HeadingField);
            this.groupBox_FieldList.Controls.Add(this.comboBox_Label2StyleField);
            this.groupBox_FieldList.Controls.Add(this.label_Label2StyleField);
            this.groupBox_FieldList.Controls.Add(this.comboBox_Label2Field);
            this.groupBox_FieldList.Controls.Add(this.label_Label2Field);
            this.groupBox_FieldList.Controls.Add(this.comboBox_Label1StyleField);
            this.groupBox_FieldList.Controls.Add(this.label_Label1StyleField);
            this.groupBox_FieldList.Controls.Add(this.comboBox_Label1Field);
            this.groupBox_FieldList.Controls.Add(this.label_Label1Field);
            this.groupBox_FieldList.Controls.Add(this.comboBox_Style2Field);
            this.groupBox_FieldList.Controls.Add(this.label_Style2Field);
            this.groupBox_FieldList.Controls.Add(this.comboBox_Style1Field);
            this.groupBox_FieldList.Controls.Add(this.label_Style1Field);
            this.groupBox_FieldList.Controls.Add(this.comboBox_ElementField);
            this.groupBox_FieldList.Controls.Add(this.label_ElementField);
            this.groupBox_FieldList.Controls.Add(this.comboBox_ColumnField);
            this.groupBox_FieldList.Controls.Add(this.label_ColumnField);
            this.groupBox_FieldList.Controls.Add(this.label_orderField);
            this.groupBox_FieldList.Controls.Add(this.comboBox_orderField);
            this.groupBox_FieldList.Name = "groupBox_FieldList";
            this.groupBox_FieldList.TabStop = false;
            // 
            // comboBox_DescriptionField
            // 
            resources.ApplyResources(this.comboBox_DescriptionField, "comboBox_DescriptionField");
            this.comboBox_DescriptionField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_DescriptionField.FormattingEnabled = true;
            this.comboBox_DescriptionField.Name = "comboBox_DescriptionField";
            this.comboBox_DescriptionField.SelectedIndexChanged += new System.EventHandler(this.comboBox_DescriptionField_SelectedIndexChanged);
            // 
            // label_DescriptionField
            // 
            resources.ApplyResources(this.label_DescriptionField, "label_DescriptionField");
            this.label_DescriptionField.Name = "label_DescriptionField";
            // 
            // comboBox_HeadingField
            // 
            resources.ApplyResources(this.comboBox_HeadingField, "comboBox_HeadingField");
            this.comboBox_HeadingField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_HeadingField.FormattingEnabled = true;
            this.comboBox_HeadingField.Name = "comboBox_HeadingField";
            this.comboBox_HeadingField.SelectedIndexChanged += new System.EventHandler(this.comboBox_HeadingField_SelectedIndexChanged);
            // 
            // label_HeadingField
            // 
            resources.ApplyResources(this.label_HeadingField, "label_HeadingField");
            this.label_HeadingField.Name = "label_HeadingField";
            // 
            // comboBox_Label2StyleField
            // 
            resources.ApplyResources(this.comboBox_Label2StyleField, "comboBox_Label2StyleField");
            this.comboBox_Label2StyleField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Label2StyleField.FormattingEnabled = true;
            this.comboBox_Label2StyleField.Name = "comboBox_Label2StyleField";
            this.comboBox_Label2StyleField.SelectedIndexChanged += new System.EventHandler(this.comboBox_Label2StyleField_SelectedIndexChanged);
            // 
            // label_Label2StyleField
            // 
            resources.ApplyResources(this.label_Label2StyleField, "label_Label2StyleField");
            this.label_Label2StyleField.Name = "label_Label2StyleField";
            // 
            // comboBox_Label2Field
            // 
            resources.ApplyResources(this.comboBox_Label2Field, "comboBox_Label2Field");
            this.comboBox_Label2Field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Label2Field.FormattingEnabled = true;
            this.comboBox_Label2Field.Name = "comboBox_Label2Field";
            this.comboBox_Label2Field.SelectedIndexChanged += new System.EventHandler(this.comboBox_Label2Field_SelectedIndexChanged);
            // 
            // label_Label2Field
            // 
            resources.ApplyResources(this.label_Label2Field, "label_Label2Field");
            this.label_Label2Field.Name = "label_Label2Field";
            // 
            // comboBox_Label1StyleField
            // 
            resources.ApplyResources(this.comboBox_Label1StyleField, "comboBox_Label1StyleField");
            this.comboBox_Label1StyleField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Label1StyleField.FormattingEnabled = true;
            this.comboBox_Label1StyleField.Name = "comboBox_Label1StyleField";
            this.comboBox_Label1StyleField.SelectedIndexChanged += new System.EventHandler(this.comboBox_Label1StyleField_SelectedIndexChanged);
            // 
            // label_Label1StyleField
            // 
            resources.ApplyResources(this.label_Label1StyleField, "label_Label1StyleField");
            this.label_Label1StyleField.Name = "label_Label1StyleField";
            // 
            // comboBox_Label1Field
            // 
            resources.ApplyResources(this.comboBox_Label1Field, "comboBox_Label1Field");
            this.comboBox_Label1Field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Label1Field.FormattingEnabled = true;
            this.comboBox_Label1Field.Name = "comboBox_Label1Field";
            this.comboBox_Label1Field.SelectedIndexChanged += new System.EventHandler(this.comboBox_Label1Field_SelectedIndexChanged);
            // 
            // label_Label1Field
            // 
            resources.ApplyResources(this.label_Label1Field, "label_Label1Field");
            this.label_Label1Field.Name = "label_Label1Field";
            // 
            // comboBox_Style2Field
            // 
            resources.ApplyResources(this.comboBox_Style2Field, "comboBox_Style2Field");
            this.comboBox_Style2Field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Style2Field.FormattingEnabled = true;
            this.comboBox_Style2Field.Name = "comboBox_Style2Field";
            this.comboBox_Style2Field.SelectedIndexChanged += new System.EventHandler(this.comboBox_Style2Field_SelectedIndexChanged);
            // 
            // label_Style2Field
            // 
            resources.ApplyResources(this.label_Style2Field, "label_Style2Field");
            this.label_Style2Field.Name = "label_Style2Field";
            // 
            // comboBox_Style1Field
            // 
            resources.ApplyResources(this.comboBox_Style1Field, "comboBox_Style1Field");
            this.comboBox_Style1Field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Style1Field.FormattingEnabled = true;
            this.comboBox_Style1Field.Name = "comboBox_Style1Field";
            this.comboBox_Style1Field.SelectedIndexChanged += new System.EventHandler(this.comboBox_Style1Field_SelectedIndexChanged);
            // 
            // label_Style1Field
            // 
            resources.ApplyResources(this.label_Style1Field, "label_Style1Field");
            this.label_Style1Field.Name = "label_Style1Field";
            // 
            // comboBox_ElementField
            // 
            resources.ApplyResources(this.comboBox_ElementField, "comboBox_ElementField");
            this.comboBox_ElementField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_ElementField.FormattingEnabled = true;
            this.comboBox_ElementField.Name = "comboBox_ElementField";
            this.comboBox_ElementField.SelectedIndexChanged += new System.EventHandler(this.comboBox_ElementField_SelectedIndexChanged);
            // 
            // label_ElementField
            // 
            resources.ApplyResources(this.label_ElementField, "label_ElementField");
            this.label_ElementField.Name = "label_ElementField";
            // 
            // comboBox_ColumnField
            // 
            resources.ApplyResources(this.comboBox_ColumnField, "comboBox_ColumnField");
            this.comboBox_ColumnField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_ColumnField.FormattingEnabled = true;
            this.comboBox_ColumnField.Name = "comboBox_ColumnField";
            this.comboBox_ColumnField.SelectedIndexChanged += new System.EventHandler(this.comboBox_ColumnField_SelectedIndexChanged);
            // 
            // label_ColumnField
            // 
            resources.ApplyResources(this.label_ColumnField, "label_ColumnField");
            this.label_ColumnField.Name = "label_ColumnField";
            // 
            // label_orderField
            // 
            resources.ApplyResources(this.label_orderField, "label_orderField");
            this.label_orderField.Name = "label_orderField";
            // 
            // comboBox_orderField
            // 
            resources.ApplyResources(this.comboBox_orderField, "comboBox_orderField");
            this.comboBox_orderField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_orderField.FormattingEnabled = true;
            this.comboBox_orderField.Name = "comboBox_orderField";
            this.comboBox_orderField.SelectedIndexChanged += new System.EventHandler(this.comboBox_orderField_SelectedIndexChanged);
            // 
            // checkBox_DEMBoxes
            // 
            resources.ApplyResources(this.checkBox_DEMBoxes, "checkBox_DEMBoxes");
            this.checkBox_DEMBoxes.Name = "checkBox_DEMBoxes";
            this.checkBox_DEMBoxes.UseVisualStyleBackColor = true;
            // 
            // comboBox_SelectTable
            // 
            resources.ApplyResources(this.comboBox_SelectTable, "comboBox_SelectTable");
            this.comboBox_SelectTable.FormattingEnabled = true;
            this.comboBox_SelectTable.Name = "comboBox_SelectTable";
            this.comboBox_SelectTable.SelectedIndexChanged += new System.EventHandler(this.comboBox_SelectTable_SelectedIndexChanged);
            // 
            // checkBox_autoCalculateColumns
            // 
            resources.ApplyResources(this.checkBox_autoCalculateColumns, "checkBox_autoCalculateColumns");
            this.checkBox_autoCalculateColumns.Name = "checkBox_autoCalculateColumns";
            this.checkBox_autoCalculateColumns.UseVisualStyleBackColor = true;
            // 
            // Form_Legend_Renderer
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.checkBox_autoCalculateColumns);
            this.Controls.Add(this.comboBox_SelectTable);
            this.Controls.Add(this.checkBox_DEMBoxes);
            this.Controls.Add(this.groupBox_FieldList);
            this.Controls.Add(this.button_Cancel);
            this.Controls.Add(this.button_Start);
            this.Controls.Add(this.button_selectTable);
            this.Controls.Add(this.label_selectTable);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form_Legend_Renderer";
            this.TopMost = true;
            this.groupBox_FieldList.ResumeLayout(false);
            this.groupBox_FieldList.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label_selectTable;
        private System.Windows.Forms.Button button_selectTable;
        private System.Windows.Forms.Button button_Start;
        private System.Windows.Forms.Button button_Cancel;
        private System.Windows.Forms.GroupBox groupBox_FieldList;
        private System.Windows.Forms.ComboBox comboBox_DescriptionField;
        private System.Windows.Forms.Label label_DescriptionField;
        private System.Windows.Forms.ComboBox comboBox_HeadingField;
        private System.Windows.Forms.Label label_HeadingField;
        private System.Windows.Forms.ComboBox comboBox_Label2StyleField;
        private System.Windows.Forms.Label label_Label2StyleField;
        private System.Windows.Forms.ComboBox comboBox_Label2Field;
        private System.Windows.Forms.Label label_Label2Field;
        private System.Windows.Forms.ComboBox comboBox_Label1StyleField;
        private System.Windows.Forms.Label label_Label1StyleField;
        private System.Windows.Forms.ComboBox comboBox_Label1Field;
        private System.Windows.Forms.Label label_Label1Field;
        private System.Windows.Forms.ComboBox comboBox_Style2Field;
        private System.Windows.Forms.Label label_Style2Field;
        private System.Windows.Forms.ComboBox comboBox_Style1Field;
        private System.Windows.Forms.Label label_Style1Field;
        private System.Windows.Forms.ComboBox comboBox_ElementField;
        private System.Windows.Forms.Label label_ElementField;
        private System.Windows.Forms.ComboBox comboBox_ColumnField;
        private System.Windows.Forms.Label label_ColumnField;
        private System.Windows.Forms.Label label_orderField;
        private System.Windows.Forms.ComboBox comboBox_orderField;
        private System.Windows.Forms.CheckBox checkBox_DEMBoxes;
        private System.Windows.Forms.ComboBox comboBox_SelectTable;
        private System.Windows.Forms.CheckBox checkBox_autoCalculateColumns;
    }
}