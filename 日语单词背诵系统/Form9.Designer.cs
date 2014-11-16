namespace 日语单词背诵系统
{
    partial class Form9
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
            this.components = new System.ComponentModel.Container();
            this.dotNetBarManager1 = new DevComponents.DotNetBar.DotNetBarManager(this.components);
            this.dockSite4 = new DevComponents.DotNetBar.DockSite();
            this.dockSite1 = new DevComponents.DotNetBar.DockSite();
            this.bar1 = new DevComponents.DotNetBar.Bar();
            this.panelDockContainer1 = new DevComponents.DotNetBar.PanelDockContainer();
            this.dockContainerItem1 = new DevComponents.DotNetBar.DockContainerItem();
            this.dockSite2 = new DevComponents.DotNetBar.DockSite();
            this.dockSite8 = new DevComponents.DotNetBar.DockSite();
            this.dockSite5 = new DevComponents.DotNetBar.DockSite();
            this.dockSite6 = new DevComponents.DotNetBar.DockSite();
            this.dockSite7 = new DevComponents.DotNetBar.DockSite();
            this.dockSite3 = new DevComponents.DotNetBar.DockSite();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.dockSite1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bar1)).BeginInit();
            this.bar1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dotNetBarManager1
            // 
            this.dotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.F1);
            this.dotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlC);
            this.dotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlA);
            this.dotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlV);
            this.dotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlX);
            this.dotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlZ);
            this.dotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlY);
            this.dotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.Del);
            this.dotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.Ins);
            this.dotNetBarManager1.BottomDockSite = this.dockSite4;
            this.dotNetBarManager1.EnableFullSizeDock = false;
            this.dotNetBarManager1.LeftDockSite = this.dockSite1;
            this.dotNetBarManager1.ParentForm = this;
            this.dotNetBarManager1.RightDockSite = this.dockSite2;
            this.dotNetBarManager1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2003;
            this.dotNetBarManager1.ToolbarBottomDockSite = this.dockSite8;
            this.dotNetBarManager1.ToolbarLeftDockSite = this.dockSite5;
            this.dotNetBarManager1.ToolbarRightDockSite = this.dockSite6;
            this.dotNetBarManager1.ToolbarTopDockSite = this.dockSite7;
            this.dotNetBarManager1.TopDockSite = this.dockSite3;
            // 
            // dockSite4
            // 
            this.dockSite4.AccessibleRole = System.Windows.Forms.AccessibleRole.Window;
            this.dockSite4.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dockSite4.DocumentDockContainer = new DevComponents.DotNetBar.DocumentDockContainer();
            this.dockSite4.Location = new System.Drawing.Point(0, 293);
            this.dockSite4.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dockSite4.Name = "dockSite4";
            this.dockSite4.Size = new System.Drawing.Size(663, 0);
            this.dockSite4.TabIndex = 3;
            this.dockSite4.TabStop = false;
            // 
            // dockSite1
            // 
            this.dockSite1.AccessibleRole = System.Windows.Forms.AccessibleRole.Window;
            this.dockSite1.Controls.Add(this.bar1);
            this.dockSite1.Dock = System.Windows.Forms.DockStyle.Left;
            this.dockSite1.DocumentDockContainer = new DevComponents.DotNetBar.DocumentDockContainer(new DevComponents.DotNetBar.DocumentBaseContainer[] {
            ((DevComponents.DotNetBar.DocumentBaseContainer)(new DevComponents.DotNetBar.DocumentBarContainer(this.bar1, 95, 293)))}, DevComponents.DotNetBar.eOrientation.Horizontal);
            this.dockSite1.Location = new System.Drawing.Point(0, 0);
            this.dockSite1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dockSite1.Name = "dockSite1";
            this.dockSite1.Size = new System.Drawing.Size(98, 293);
            this.dockSite1.TabIndex = 0;
            this.dockSite1.TabStop = false;
            // 
            // bar1
            // 
            this.bar1.AccessibleDescription = "DotNetBar Bar (bar1)";
            this.bar1.AccessibleName = "DotNetBar Bar";
            this.bar1.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar;
            this.bar1.AutoSyncBarCaption = true;
            this.bar1.CloseSingleTab = true;
            this.bar1.Controls.Add(this.panelDockContainer1);
            this.bar1.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption;
            this.bar1.Items.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.dockContainerItem1});
            this.bar1.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer;
            this.bar1.Location = new System.Drawing.Point(0, 0);
            this.bar1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.bar1.Name = "bar1";
            this.bar1.Size = new System.Drawing.Size(95, 293);
            this.bar1.Stretch = true;
            this.bar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2003;
            this.bar1.TabIndex = 0;
            this.bar1.TabStop = false;
            this.bar1.Text = "dockContainerItem1";
            // 
            // panelDockContainer1
            // 
            this.panelDockContainer1.Location = new System.Drawing.Point(3, 23);
            this.panelDockContainer1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelDockContainer1.Name = "panelDockContainer1";
            this.panelDockContainer1.Size = new System.Drawing.Size(89, 267);
            this.panelDockContainer1.Style.Alignment = System.Drawing.StringAlignment.Center;
            this.panelDockContainer1.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.panelDockContainer1.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.panelDockContainer1.Style.Border = DevComponents.DotNetBar.eBorderType.SingleLine;
            this.panelDockContainer1.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.panelDockContainer1.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.panelDockContainer1.Style.GradientAngle = 90;
            this.panelDockContainer1.TabIndex = 0;
            // 
            // dockContainerItem1
            // 
            this.dockContainerItem1.Control = this.panelDockContainer1;
            this.dockContainerItem1.Name = "dockContainerItem1";
            this.dockContainerItem1.Text = "dockContainerItem1";
            // 
            // dockSite2
            // 
            this.dockSite2.AccessibleRole = System.Windows.Forms.AccessibleRole.Window;
            this.dockSite2.Dock = System.Windows.Forms.DockStyle.Right;
            this.dockSite2.DocumentDockContainer = new DevComponents.DotNetBar.DocumentDockContainer();
            this.dockSite2.Location = new System.Drawing.Point(663, 0);
            this.dockSite2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dockSite2.Name = "dockSite2";
            this.dockSite2.Size = new System.Drawing.Size(0, 293);
            this.dockSite2.TabIndex = 1;
            this.dockSite2.TabStop = false;
            // 
            // dockSite8
            // 
            this.dockSite8.AccessibleRole = System.Windows.Forms.AccessibleRole.Window;
            this.dockSite8.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dockSite8.Location = new System.Drawing.Point(0, 293);
            this.dockSite8.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dockSite8.Name = "dockSite8";
            this.dockSite8.Size = new System.Drawing.Size(663, 0);
            this.dockSite8.TabIndex = 7;
            this.dockSite8.TabStop = false;
            // 
            // dockSite5
            // 
            this.dockSite5.AccessibleRole = System.Windows.Forms.AccessibleRole.Window;
            this.dockSite5.Dock = System.Windows.Forms.DockStyle.Left;
            this.dockSite5.Location = new System.Drawing.Point(0, 0);
            this.dockSite5.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dockSite5.Name = "dockSite5";
            this.dockSite5.Size = new System.Drawing.Size(0, 293);
            this.dockSite5.TabIndex = 4;
            this.dockSite5.TabStop = false;
            // 
            // dockSite6
            // 
            this.dockSite6.AccessibleRole = System.Windows.Forms.AccessibleRole.Window;
            this.dockSite6.Dock = System.Windows.Forms.DockStyle.Right;
            this.dockSite6.Location = new System.Drawing.Point(663, 0);
            this.dockSite6.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dockSite6.Name = "dockSite6";
            this.dockSite6.Size = new System.Drawing.Size(0, 293);
            this.dockSite6.TabIndex = 5;
            this.dockSite6.TabStop = false;
            // 
            // dockSite7
            // 
            this.dockSite7.AccessibleRole = System.Windows.Forms.AccessibleRole.Window;
            this.dockSite7.Dock = System.Windows.Forms.DockStyle.Top;
            this.dockSite7.Location = new System.Drawing.Point(0, 0);
            this.dockSite7.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dockSite7.Name = "dockSite7";
            this.dockSite7.Size = new System.Drawing.Size(663, 0);
            this.dockSite7.TabIndex = 6;
            this.dockSite7.TabStop = false;
            // 
            // dockSite3
            // 
            this.dockSite3.AccessibleRole = System.Windows.Forms.AccessibleRole.Window;
            this.dockSite3.Dock = System.Windows.Forms.DockStyle.Top;
            this.dockSite3.DocumentDockContainer = new DevComponents.DotNetBar.DocumentDockContainer();
            this.dockSite3.Location = new System.Drawing.Point(0, 0);
            this.dockSite3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dockSite3.Name = "dockSite3";
            this.dockSite3.Size = new System.Drawing.Size(663, 0);
            this.dockSite3.TabIndex = 2;
            this.dockSite3.TabStop = false;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("微软雅黑", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(258, 64);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(231, 99);
            this.button1.TabIndex = 8;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(258, 204);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(98, 35);
            this.button2.TabIndex = 9;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(391, 204);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(98, 35);
            this.button3.TabIndex = 10;
            this.button3.Text = "button3";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // Form9
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(663, 293);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dockSite2);
            this.Controls.Add(this.dockSite1);
            this.Controls.Add(this.dockSite3);
            this.Controls.Add(this.dockSite4);
            this.Controls.Add(this.dockSite5);
            this.Controls.Add(this.dockSite6);
            this.Controls.Add(this.dockSite7);
            this.Controls.Add(this.dockSite8);
            this.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form9";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form9";
            this.Load += new System.EventHandler(this.Form9_Load);
            this.dockSite1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.bar1)).EndInit();
            this.bar1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.DotNetBarManager dotNetBarManager1;
        private DevComponents.DotNetBar.DockSite dockSite4;
        private DevComponents.DotNetBar.DockSite dockSite1;
        private DevComponents.DotNetBar.Bar bar1;
        private DevComponents.DotNetBar.PanelDockContainer panelDockContainer1;
        private DevComponents.DotNetBar.DockContainerItem dockContainerItem1;
        private DevComponents.DotNetBar.DockSite dockSite2;
        private DevComponents.DotNetBar.DockSite dockSite3;
        private DevComponents.DotNetBar.DockSite dockSite5;
        private DevComponents.DotNetBar.DockSite dockSite6;
        private DevComponents.DotNetBar.DockSite dockSite7;
        private DevComponents.DotNetBar.DockSite dockSite8;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
    }
}