namespace FCMF0611
{
    partial class FCMF0611_PRINT
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo oraConnectionInfo1 = new InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement1 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement5 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement3 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement4 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement6 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement2 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.isMessageAdapter1 = new InfoSummit.Win.ControlAdv.ISMessageAdapter(this.components);
            this.btnPRINT = new InfoSummit.Win.ControlAdv.ISButton();
            this.igbCONFIRM_INFOMATION = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.ACCOUNT = new InfoSummit.Win.ControlAdv.ISRadioButtonAdv();
            this.DEPARTMENT = new InfoSummit.Win.ControlAdv.ISRadioButtonAdv();
            this.btnCANCEL = new InfoSummit.Win.ControlAdv.ISButton();
            this.PRINT_TYPE = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.igbCONFIRM_INFOMATION)).BeginInit();
            this.igbCONFIRM_INFOMATION.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ACCOUNT)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DEPARTMENT)).BeginInit();
            // 
            // isAppInterfaceAdv1
            // 
            this.isAppInterfaceAdv1.AppMainButtonClick += new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv.ButtonEventHandler(this.isAppInterfaceAdv1_AppMainButtonClick);
            // 
            // isOraConnection1
            // 
            this.isOraConnection1.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isOraConnection1.OraConnectionInfo = oraConnectionInfo1;
            this.isOraConnection1.OraHost = "211.168.59.26";
            this.isOraConnection1.OraPassword = "infoflex";
            this.isOraConnection1.OraPort = "1521";
            this.isOraConnection1.OraServiceName = "FXCDB";
            this.isOraConnection1.OraUserId = "APPS";
            // 
            // isMessageAdapter1
            // 
            this.isMessageAdapter1.OraConnection = this.isOraConnection1;
            // 
            // btnPRINT
            // 
            this.btnPRINT.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.btnPRINT.ButtonText = "Printing";
            isLanguageElement1.Default = "Printing";
            isLanguageElement1.SiteName = null;
            isLanguageElement1.TL1_KR = "인쇄";
            isLanguageElement1.TL2_CN = "";
            isLanguageElement1.TL3_VN = "";
            isLanguageElement1.TL4_JP = "";
            isLanguageElement1.TL5_XAA = "";
            this.btnPRINT.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement1});
            // 
            // FCMF0611_PRINT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(278, 137);
            this.ControlBox = false;
            this.Controls.Add(this.PRINT_TYPE);
            this.Controls.Add(this.igbCONFIRM_INFOMATION);
            this.Controls.Add(this.btnCANCEL);
            this.Controls.Add(this.btnPRINT);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FCMF0611_PRINT";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Set Printing";
            this.Load += new System.EventHandler(this.FCMF0611_FILE_Load);
            this.Shown += new System.EventHandler(this.FCMF0611_FILE_Shown);
            this.btnPRINT.Location = new System.Drawing.Point(94, 98);
            this.btnPRINT.Name = "btnPRINT";
            this.btnPRINT.Size = new System.Drawing.Size(75, 25);
            this.btnPRINT.TabIndex = 0;
            this.btnPRINT.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.btnPRINT_ButtonClick);
            // 
            // igbCONFIRM_INFOMATION
            // 
            this.igbCONFIRM_INFOMATION.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.igbCONFIRM_INFOMATION.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.igbCONFIRM_INFOMATION.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.igbCONFIRM_INFOMATION.Controls.Add(this.ACCOUNT);
            this.igbCONFIRM_INFOMATION.Controls.Add(this.DEPARTMENT);
            this.igbCONFIRM_INFOMATION.Location = new System.Drawing.Point(8, 8);
            this.igbCONFIRM_INFOMATION.Name = "igbCONFIRM_INFOMATION";
            this.igbCONFIRM_INFOMATION.PromptText = "Confirm Infomation";
            isLanguageElement5.Default = "Confirm Infomation";
            isLanguageElement5.SiteName = null;
            isLanguageElement5.TL1_KR = "승인 정보";
            isLanguageElement5.TL2_CN = "";
            isLanguageElement5.TL3_VN = "";
            isLanguageElement5.TL4_JP = "";
            isLanguageElement5.TL5_XAA = "";
            this.igbCONFIRM_INFOMATION.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement5});
            this.igbCONFIRM_INFOMATION.PromptVisible = false;
            this.igbCONFIRM_INFOMATION.Size = new System.Drawing.Size(262, 76);
            this.igbCONFIRM_INFOMATION.TabIndex = 0;
            // 
            // ACCOUNT
            // 
            this.ACCOUNT.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.ACCOUNT.DataAdapter = null;
            this.ACCOUNT.DataColumn = null;
            this.ACCOUNT.Location = new System.Drawing.Point(44, 40);
            this.ACCOUNT.MetroColor = System.Drawing.Color.Empty;
            this.ACCOUNT.Name = "ACCOUNT";
            this.ACCOUNT.Office2007ColorScheme = Syncfusion.Windows.Forms.Office2007Theme.Managed;
            this.ACCOUNT.PromptText = "예산신청내역(계정별)";
            isLanguageElement3.Default = "Budget request list(Account)";
            isLanguageElement3.SiteName = null;
            isLanguageElement3.TL1_KR = "예산신청내역(계정별)";
            isLanguageElement3.TL2_CN = null;
            isLanguageElement3.TL3_VN = null;
            isLanguageElement3.TL4_JP = null;
            isLanguageElement3.TL5_XAA = null;
            this.ACCOUNT.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement3});
            this.ACCOUNT.RadioButtonValue = null;
            this.ACCOUNT.RadioCheckedString = "A";
            this.ACCOUNT.Size = new System.Drawing.Size(210, 24);
            this.ACCOUNT.Style = Syncfusion.Windows.Forms.Tools.RadioButtonAdvStyle.Office2007;
            this.ACCOUNT.TabIndex = 1;
            this.ACCOUNT.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.ACCOUNT.Text = "예산신청내역(계정별)";
            this.ACCOUNT.ThemesEnabled = false;
            this.ACCOUNT.CheckChanged += new System.EventHandler(this.RB_SUMMARY_CheckChanged);
            // 
            // DEPARTMENT
            // 
            this.DEPARTMENT.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.DEPARTMENT.CheckedString = "H";
            this.DEPARTMENT.DataAdapter = null;
            this.DEPARTMENT.DataColumn = null;
            this.DEPARTMENT.Location = new System.Drawing.Point(44, 10);
            this.DEPARTMENT.MetroColor = System.Drawing.Color.Empty;
            this.DEPARTMENT.Name = "DEPARTMENT";
            this.DEPARTMENT.Office2007ColorScheme = Syncfusion.Windows.Forms.Office2007Theme.Managed;
            this.DEPARTMENT.PromptText = "예산신청내역(부서별)";
            isLanguageElement4.Default = "Budget request list(Department)";
            isLanguageElement4.SiteName = null;
            isLanguageElement4.TL1_KR = "예산신청내역(부서별)";
            isLanguageElement4.TL2_CN = null;
            isLanguageElement4.TL3_VN = null;
            isLanguageElement4.TL4_JP = null;
            isLanguageElement4.TL5_XAA = null;
            this.DEPARTMENT.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement4});
            this.DEPARTMENT.RadioButtonValue = null;
            this.DEPARTMENT.RadioCheckedString = "D";
            this.DEPARTMENT.Size = new System.Drawing.Size(210, 24);
            this.DEPARTMENT.Style = Syncfusion.Windows.Forms.Tools.RadioButtonAdvStyle.Office2007;
            this.DEPARTMENT.TabIndex = 0;
            this.DEPARTMENT.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.DEPARTMENT.Text = "예산신청내역(부서별)";
            this.DEPARTMENT.ThemesEnabled = false;
            this.DEPARTMENT.UncheckedString = "1";
            this.DEPARTMENT.CheckChanged += new System.EventHandler(this.RB_SUMMARY_CheckChanged);
            // 
            // btnCANCEL
            // 
            this.btnCANCEL.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.btnCANCEL.ButtonText = "Cancel";
            isLanguageElement6.Default = "Cancel";
            isLanguageElement6.SiteName = null;
            isLanguageElement6.TL1_KR = "취소";
            isLanguageElement6.TL2_CN = "";
            isLanguageElement6.TL3_VN = "";
            isLanguageElement6.TL4_JP = "";
            isLanguageElement6.TL5_XAA = "";
            this.btnCANCEL.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement6});
            this.btnCANCEL.Location = new System.Drawing.Point(175, 98);
            this.btnCANCEL.Name = "btnCANCEL";
            this.btnCANCEL.Size = new System.Drawing.Size(75, 25);
            this.btnCANCEL.TabIndex = 1;
            this.btnCANCEL.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.btnCANCEL_ButtonClick);
            // 
            // PRINT_TYPE
            // 
            this.PRINT_TYPE.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.PRINT_TYPE.ComboBoxValue = "";
            this.PRINT_TYPE.ComboData = null;
            this.PRINT_TYPE.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.PRINT_TYPE.DataAdapter = null;
            this.PRINT_TYPE.DataColumn = null;
            this.PRINT_TYPE.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.PRINT_TYPE.DoubleValue = 0;
            this.PRINT_TYPE.EditValue = "";
            this.PRINT_TYPE.Insertable = false;
            this.PRINT_TYPE.Location = new System.Drawing.Point(11, 90);
            this.PRINT_TYPE.LookupAdapter = null;
            this.PRINT_TYPE.Name = "PRINT_TYPE";
            this.PRINT_TYPE.Nullable = true;
            this.PRINT_TYPE.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.PRINT_TYPE.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.PRINT_TYPE.PromptText = "Print type";
            isLanguageElement2.Default = "Print type";
            isLanguageElement2.SiteName = null;
            isLanguageElement2.TL1_KR = "인쇄구분";
            isLanguageElement2.TL2_CN = null;
            isLanguageElement2.TL3_VN = null;
            isLanguageElement2.TL4_JP = null;
            isLanguageElement2.TL5_XAA = null;
            this.PRINT_TYPE.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement2});
            this.PRINT_TYPE.PromptVisible = false;
            this.PRINT_TYPE.ReadOnly = true;
            this.PRINT_TYPE.Size = new System.Drawing.Size(24, 21);
            this.PRINT_TYPE.TabIndex = 3;
            this.PRINT_TYPE.TabStop = false;
            this.PRINT_TYPE.TextValue = "";
            this.PRINT_TYPE.Updatable = false;
            this.PRINT_TYPE.Visible = false;
            this.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.igbCONFIRM_INFOMATION)).EndInit();
            this.igbCONFIRM_INFOMATION.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ACCOUNT)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DEPARTMENT)).EndInit();

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISOraConnection isOraConnection1;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter isMessageAdapter1;
        private InfoSummit.Win.ControlAdv.ISButton btnPRINT;
        private InfoSummit.Win.ControlAdv.ISGroupBox igbCONFIRM_INFOMATION;
        private InfoSummit.Win.ControlAdv.ISButton btnCANCEL;
        private InfoSummit.Win.ControlAdv.ISEditAdv PRINT_TYPE;
        private InfoSummit.Win.ControlAdv.ISRadioButtonAdv ACCOUNT;
        private InfoSummit.Win.ControlAdv.ISRadioButtonAdv DEPARTMENT;
    }
}

