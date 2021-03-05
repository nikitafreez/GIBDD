namespace GIBDD
{
    partial class Authorization
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.CloseButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.loginBox = new System.Windows.Forms.TextBox();
            this.PasswordBox = new System.Windows.Forms.TextBox();
            this.auth_button = new System.Windows.Forms.Button();
            this.reg_button = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(297, 75);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(178, 31);
            this.label1.TabIndex = 0;
            this.label1.Text = "Авторизация";
            // 
            // CloseButton
            // 
            this.CloseButton.BackColor = System.Drawing.Color.Red;
            this.CloseButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.CloseButton.Location = new System.Drawing.Point(598, 74);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(113, 32);
            this.CloseButton.TabIndex = 1;
            this.CloseButton.Text = "Закрыть";
            this.CloseButton.UseVisualStyleBackColor = false;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(172, 181);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 17);
            this.label2.TabIndex = 2;
            this.label2.Text = "Логин";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(172, 220);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 17);
            this.label3.TabIndex = 3;
            this.label3.Text = "Пароль";
            // 
            // loginBox
            // 
            this.loginBox.Location = new System.Drawing.Point(269, 181);
            this.loginBox.Name = "loginBox";
            this.loginBox.Size = new System.Drawing.Size(241, 22);
            this.loginBox.TabIndex = 4;
            // 
            // PasswordBox
            // 
            this.PasswordBox.Location = new System.Drawing.Point(269, 220);
            this.PasswordBox.Name = "PasswordBox";
            this.PasswordBox.Size = new System.Drawing.Size(241, 22);
            this.PasswordBox.TabIndex = 5;
            // 
            // auth_button
            // 
            this.auth_button.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.auth_button.Location = new System.Drawing.Point(303, 260);
            this.auth_button.Name = "auth_button";
            this.auth_button.Size = new System.Drawing.Size(172, 32);
            this.auth_button.TabIndex = 6;
            this.auth_button.Text = "Авторизация";
            this.auth_button.UseVisualStyleBackColor = false;
            this.auth_button.Click += new System.EventHandler(this.button1_Click);
            // 
            // reg_button
            // 
            this.reg_button.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.reg_button.Location = new System.Drawing.Point(303, 307);
            this.reg_button.Name = "reg_button";
            this.reg_button.Size = new System.Drawing.Size(172, 32);
            this.reg_button.TabIndex = 8;
            this.reg_button.Text = "Регистрация";
            this.reg_button.UseVisualStyleBackColor = false;
            this.reg_button.Click += new System.EventHandler(this.button3_Click);
            // 
            // Authorization
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.reg_button);
            this.Controls.Add(this.auth_button);
            this.Controls.Add(this.PasswordBox);
            this.Controls.Add(this.loginBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.CloseButton);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Authorization";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Авторизация";
            this.Load += new System.EventHandler(this.Authorization_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button CloseButton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox loginBox;
        private System.Windows.Forms.TextBox PasswordBox;
        private System.Windows.Forms.Button auth_button;
        private System.Windows.Forms.Button reg_button;
    }
}