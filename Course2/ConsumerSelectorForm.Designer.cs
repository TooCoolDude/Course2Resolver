
namespace Course2
{
    partial class ConsumerSelectorForm
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
            comboBox1 = new ComboBox();
            button1 = new Button();
            SuspendLayout();
            // 
            // comboBox1
            // 
            comboBox1.FormattingEnabled = true;
            comboBox1.Location = new Point(135, 53);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(502, 40);
            comboBox1.TabIndex = 0;
            // 
            // button1
            // 
            button1.Location = new Point(303, 217);
            button1.Name = "button1";
            button1.Size = new Size(150, 46);
            button1.TabIndex = 1;
            button1.Text = "select";
            button1.UseVisualStyleBackColor = true;
            button1.Click += this.button1_Click;
            // 
            // ConsumerSelectorForm
            // 
            AutoScaleDimensions = new SizeF(13F, 32F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(button1);
            Controls.Add(comboBox1);
            Name = "ConsumerSelectorForm";
            Text = "CansumerSelectorForm";
            ResumeLayout(false);
        }

        #endregion

        private ComboBox comboBox1;
        private Button button1;
    }
}