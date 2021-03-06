﻿namespace SlideShowImageFinder
{
    partial class ImageGenerationPage
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
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.nextButton = new System.Windows.Forms.Button();
            this.makeSlideButton = new System.Windows.Forms.Button();
            this.clearPicturesButton = new System.Windows.Forms.Button();
            this.previousImagesButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(12, 59);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(241, 270);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            this.pictureBox1.DoubleClick += new System.EventHandler(this.pictureBox1_DoubleClick);
            // 
            // pictureBox2
            // 
            this.pictureBox2.Location = new System.Drawing.Point(279, 59);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(241, 270);
            this.pictureBox2.TabIndex = 1;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click);
            this.pictureBox2.DoubleClick += new System.EventHandler(this.pictureBox2_DoubleClick);
            // 
            // pictureBox3
            // 
            this.pictureBox3.Location = new System.Drawing.Point(542, 59);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(246, 270);
            this.pictureBox3.TabIndex = 2;
            this.pictureBox3.TabStop = false;
            this.pictureBox3.Click += new System.EventHandler(this.pictureBox3_Click);
            this.pictureBox3.DoubleClick += new System.EventHandler(this.pictureBox3_DoubleClick);
            // 
            // nextButton
            // 
            this.nextButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nextButton.Location = new System.Drawing.Point(650, 370);
            this.nextButton.Name = "nextButton";
            this.nextButton.Size = new System.Drawing.Size(138, 59);
            this.nextButton.TabIndex = 10;
            this.nextButton.Text = "Next Images";
            this.nextButton.UseVisualStyleBackColor = true;
            this.nextButton.Click += new System.EventHandler(this.nextButton_Click);
            // 
            // makeSlideButton
            // 
            this.makeSlideButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.makeSlideButton.Location = new System.Drawing.Point(12, 370);
            this.makeSlideButton.Name = "makeSlideButton";
            this.makeSlideButton.Size = new System.Drawing.Size(125, 59);
            this.makeSlideButton.TabIndex = 11;
            this.makeSlideButton.Text = "Make Slide";
            this.makeSlideButton.UseVisualStyleBackColor = true;
            this.makeSlideButton.Click += new System.EventHandler(this.makeSlideButton_Click);
            // 
            // clearPicturesButton
            // 
            this.clearPicturesButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clearPicturesButton.Location = new System.Drawing.Point(211, 370);
            this.clearPicturesButton.Name = "clearPicturesButton";
            this.clearPicturesButton.Size = new System.Drawing.Size(160, 59);
            this.clearPicturesButton.TabIndex = 12;
            this.clearPicturesButton.Text = "Clear Pictures";
            this.clearPicturesButton.UseVisualStyleBackColor = true;
            this.clearPicturesButton.Click += new System.EventHandler(this.clearPicturesButton_Click);
            // 
            // previousImagesButton
            // 
            this.previousImagesButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.previousImagesButton.Location = new System.Drawing.Point(431, 370);
            this.previousImagesButton.Name = "previousImagesButton";
            this.previousImagesButton.Size = new System.Drawing.Size(157, 59);
            this.previousImagesButton.TabIndex = 13;
            this.previousImagesButton.Text = "Previous Images";
            this.previousImagesButton.UseVisualStyleBackColor = true;
            this.previousImagesButton.Click += new System.EventHandler(this.previousImagesButton_Click);
            // 
            // ImageGenerationPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.previousImagesButton);
            this.Controls.Add(this.clearPicturesButton);
            this.Controls.Add(this.makeSlideButton);
            this.Controls.Add(this.nextButton);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox1);
            this.Name = "ImageGenerationPage";
            this.Text = "Image Generation Page";
            this.Load += new System.EventHandler(this.ImageGenerationPage_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.Button nextButton;
        private System.Windows.Forms.Button makeSlideButton;
        private System.Windows.Forms.Button clearPicturesButton;
        private System.Windows.Forms.Button previousImagesButton;
    }
}