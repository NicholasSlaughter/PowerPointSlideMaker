namespace SlideShowImageFinder
{
    partial class textTextbox
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
            this.titleLabel = new System.Windows.Forms.Label();
            this.textLabel = new System.Windows.Forms.Label();
            this.titleTextBox = new System.Windows.Forms.TextBox();
            this.generateButton = new System.Windows.Forms.Button();
            this.appTitle = new System.Windows.Forms.Label();
            this.richTextBoxText = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // titleLabel
            // 
            this.titleLabel.AutoSize = true;
            this.titleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 30F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titleLabel.Location = new System.Drawing.Point(150, 91);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(106, 46);
            this.titleLabel.TabIndex = 0;
            this.titleLabel.Text = "Title:";
            // 
            // textLabel
            // 
            this.textLabel.AutoSize = true;
            this.textLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 30F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textLabel.Location = new System.Drawing.Point(150, 152);
            this.textLabel.Name = "textLabel";
            this.textLabel.Size = new System.Drawing.Size(109, 46);
            this.textLabel.TabIndex = 1;
            this.textLabel.Text = "Text:";
            // 
            // titleTextBox
            // 
            this.titleTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 30F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titleTextBox.Location = new System.Drawing.Point(323, 88);
            this.titleTextBox.Name = "titleTextBox";
            this.titleTextBox.Size = new System.Drawing.Size(397, 53);
            this.titleTextBox.TabIndex = 2;
            // 
            // generateButton
            // 
            this.generateButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generateButton.Location = new System.Drawing.Point(323, 385);
            this.generateButton.Name = "generateButton";
            this.generateButton.Size = new System.Drawing.Size(110, 53);
            this.generateButton.TabIndex = 4;
            this.generateButton.Text = "Generate";
            this.generateButton.UseVisualStyleBackColor = true;
            this.generateButton.Click += new System.EventHandler(this.generateButton_Click);
            // 
            // appTitle
            // 
            this.appTitle.AutoSize = true;
            this.appTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 30F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.appTitle.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.appTitle.Location = new System.Drawing.Point(169, 9);
            this.appTitle.Name = "appTitle";
            this.appTitle.Size = new System.Drawing.Size(464, 46);
            this.appTitle.TabIndex = 5;
            this.appTitle.Text = "Slide Show Image Finder";
            this.appTitle.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // richTextBoxText
            // 
            this.richTextBoxText.Location = new System.Drawing.Point(323, 152);
            this.richTextBoxText.Name = "richTextBoxText";
            this.richTextBoxText.Size = new System.Drawing.Size(397, 214);
            this.richTextBoxText.TabIndex = 6;
            this.richTextBoxText.Text = "";
            // 
            // textTextbox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.richTextBoxText);
            this.Controls.Add(this.appTitle);
            this.Controls.Add(this.generateButton);
            this.Controls.Add(this.titleTextBox);
            this.Controls.Add(this.textLabel);
            this.Controls.Add(this.titleLabel);
            this.Name = "textTextbox";
            this.Text = "Slide Show Image Finder";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label titleLabel;
        private System.Windows.Forms.Label textLabel;
        private System.Windows.Forms.TextBox titleTextBox;
        private System.Windows.Forms.Button generateButton;
        private System.Windows.Forms.Label appTitle;
        private System.Windows.Forms.RichTextBox richTextBoxText;
    }
}

