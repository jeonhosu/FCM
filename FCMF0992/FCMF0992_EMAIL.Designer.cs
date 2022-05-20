namespace FCMF0992
{
    partial class FCMF0992_EMAIL
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
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement7 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement4 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement3 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement2 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo oraConnectionInfo1 = new InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement6 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement5 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement1 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement2 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement3 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement4 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement5 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement6 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement7 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement8 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement9 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement1 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            this.TAX_BILL_ISSUE_NO = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.SELL_USER_EMAIL = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.BUY_USER_EMAIL = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.BUY_USER2_EMAIL = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.isMessageAdapter1 = new InfoSummit.Win.ControlAdv.ISMessageAdapter(this.components);
            this.BTN_CLOSED = new InfoSummit.Win.ControlAdv.ISButton();
            this.isGroupBox3 = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.IDC_SET_RESEND_EMAIL = new InfoSummit.Win.ControlAdv.ISDataCommand(this.components);
            this.BTN_RESEND_EMAIL = new InfoSummit.Win.ControlAdv.ISButton();
            this.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox3)).BeginInit();
            this.isGroupBox3.SuspendLayout();
            // 
            // TAX_BILL_ISSUE_NO
            // 
            this.TAX_BILL_ISSUE_NO.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.TAX_BILL_ISSUE_NO.ComboBoxValue = "";
            this.TAX_BILL_ISSUE_NO.ComboData = null;
            this.TAX_BILL_ISSUE_NO.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.TAX_BILL_ISSUE_NO.DataAdapter = null;
            this.TAX_BILL_ISSUE_NO.DataColumn = null;
            this.TAX_BILL_ISSUE_NO.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.TAX_BILL_ISSUE_NO.DoubleValue = 0;
            this.TAX_BILL_ISSUE_NO.EditValue = "";
            // 
            // FCMF0992_EMAIL
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(413, 152);
            this.ControlBox = false;
            this.Controls.Add(this.BTN_RESEND_EMAIL);
            this.Controls.Add(this.isGroupBox3);
            this.Controls.Add(this.BTN_CLOSED);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FCMF0992_EMAIL";
            this.Padding = new System.Windows.Forms.Padding(2);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "◈Email 재발송";
            this.Load += new System.EventHandler(this.FCMF0992_EMAIL_Load);
            this.TAX_BILL_ISSUE_NO.Insertable = false;
            this.TAX_BILL_ISSUE_NO.Location = new System.Drawing.Point(3, 9);
            this.TAX_BILL_ISSUE_NO.LookupAdapter = null;
            this.TAX_BILL_ISSUE_NO.Name = "TAX_BILL_ISSUE_NO";
            this.TAX_BILL_ISSUE_NO.Nullable = true;
            this.TAX_BILL_ISSUE_NO.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.TAX_BILL_ISSUE_NO.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.TAX_BILL_ISSUE_NO.PromptText = "세금계산서 발행번호";
            isLanguageElement7.Default = "Tax bill issue NO";
            isLanguageElement7.SiteName = null;
            isLanguageElement7.TL1_KR = "세금계산서 발행번호";
            isLanguageElement7.TL2_CN = null;
            isLanguageElement7.TL3_VN = null;
            isLanguageElement7.TL4_JP = null;
            isLanguageElement7.TL5_XAA = null;
            this.TAX_BILL_ISSUE_NO.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement7});
            this.TAX_BILL_ISSUE_NO.PromptWidth = 160;
            this.TAX_BILL_ISSUE_NO.ReadOnly = true;
            this.TAX_BILL_ISSUE_NO.Size = new System.Drawing.Size(390, 21);
            this.TAX_BILL_ISSUE_NO.TabIndex = 0;
            this.TAX_BILL_ISSUE_NO.TabStop = false;
            this.TAX_BILL_ISSUE_NO.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.TAX_BILL_ISSUE_NO.TextValue = "";
            this.TAX_BILL_ISSUE_NO.Updatable = false;
            // 
            // isAppInterfaceAdv1
            // 
            this.isAppInterfaceAdv1.AppMainButtonClick += new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv.ButtonEventHandler(this.isAppInterfaceAdv1_AppMainButtonClick);
            // 
            // SELL_USER_EMAIL
            // 
            this.SELL_USER_EMAIL.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.SELL_USER_EMAIL.ComboBoxValue = "";
            this.SELL_USER_EMAIL.ComboData = null;
            this.SELL_USER_EMAIL.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.SELL_USER_EMAIL.DataAdapter = null;
            this.SELL_USER_EMAIL.DataColumn = "";
            this.SELL_USER_EMAIL.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.SELL_USER_EMAIL.DoubleValue = 0;
            this.SELL_USER_EMAIL.EditValue = "";
            this.SELL_USER_EMAIL.Insertable = false;
            this.SELL_USER_EMAIL.Location = new System.Drawing.Point(3, 34);
            this.SELL_USER_EMAIL.LookupAdapter = null;
            this.SELL_USER_EMAIL.Name = "SELL_USER_EMAIL";
            this.SELL_USER_EMAIL.Nullable = true;
            this.SELL_USER_EMAIL.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.SELL_USER_EMAIL.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.SELL_USER_EMAIL.PromptText = "공급자 이메일주소";
            isLanguageElement4.Default = "Sell User Email";
            isLanguageElement4.SiteName = null;
            isLanguageElement4.TL1_KR = "공급자 이메일주소";
            isLanguageElement4.TL2_CN = null;
            isLanguageElement4.TL3_VN = null;
            isLanguageElement4.TL4_JP = null;
            isLanguageElement4.TL5_XAA = null;
            this.SELL_USER_EMAIL.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement4});
            this.SELL_USER_EMAIL.PromptWidth = 160;
            this.SELL_USER_EMAIL.Size = new System.Drawing.Size(390, 21);
            this.SELL_USER_EMAIL.TabIndex = 1;
            this.SELL_USER_EMAIL.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.SELL_USER_EMAIL.TextValue = "";
            this.SELL_USER_EMAIL.Updatable = false;
            // 
            // BUY_USER_EMAIL
            // 
            this.BUY_USER_EMAIL.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BUY_USER_EMAIL.ComboBoxValue = "";
            this.BUY_USER_EMAIL.ComboData = null;
            this.BUY_USER_EMAIL.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.BUY_USER_EMAIL.DataAdapter = null;
            this.BUY_USER_EMAIL.DataColumn = "";
            this.BUY_USER_EMAIL.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.BUY_USER_EMAIL.DoubleValue = 0;
            this.BUY_USER_EMAIL.EditValue = "";
            this.BUY_USER_EMAIL.Insertable = false;
            this.BUY_USER_EMAIL.Location = new System.Drawing.Point(3, 61);
            this.BUY_USER_EMAIL.LookupAdapter = null;
            this.BUY_USER_EMAIL.Name = "BUY_USER_EMAIL";
            this.BUY_USER_EMAIL.Nullable = true;
            this.BUY_USER_EMAIL.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.BUY_USER_EMAIL.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.BUY_USER_EMAIL.PromptText = "공급받는자 이메일";
            isLanguageElement3.Default = "Buy User Email";
            isLanguageElement3.SiteName = null;
            isLanguageElement3.TL1_KR = "공급받는자 이메일";
            isLanguageElement3.TL2_CN = null;
            isLanguageElement3.TL3_VN = null;
            isLanguageElement3.TL4_JP = null;
            isLanguageElement3.TL5_XAA = null;
            this.BUY_USER_EMAIL.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement3});
            this.BUY_USER_EMAIL.PromptWidth = 160;
            this.BUY_USER_EMAIL.Size = new System.Drawing.Size(390, 21);
            this.BUY_USER_EMAIL.TabIndex = 2;
            this.BUY_USER_EMAIL.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.BUY_USER_EMAIL.TextValue = "";
            this.BUY_USER_EMAIL.Updatable = false;
            // 
            // BUY_USER2_EMAIL
            // 
            this.BUY_USER2_EMAIL.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BUY_USER2_EMAIL.ComboBoxValue = "";
            this.BUY_USER2_EMAIL.ComboData = null;
            this.BUY_USER2_EMAIL.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.BUY_USER2_EMAIL.DataAdapter = null;
            this.BUY_USER2_EMAIL.DataColumn = "";
            this.BUY_USER2_EMAIL.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.BUY_USER2_EMAIL.DoubleValue = 0;
            this.BUY_USER2_EMAIL.EditValue = "";
            this.BUY_USER2_EMAIL.Insertable = false;
            this.BUY_USER2_EMAIL.Location = new System.Drawing.Point(3, 84);
            this.BUY_USER2_EMAIL.LookupAdapter = null;
            this.BUY_USER2_EMAIL.Name = "BUY_USER2_EMAIL";
            this.BUY_USER2_EMAIL.Nullable = true;
            this.BUY_USER2_EMAIL.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.BUY_USER2_EMAIL.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.BUY_USER2_EMAIL.PromptText = "공급받는자 이메일2";
            isLanguageElement2.Default = "Buy User Email2";
            isLanguageElement2.SiteName = null;
            isLanguageElement2.TL1_KR = "공급받는자 이메일2";
            isLanguageElement2.TL2_CN = null;
            isLanguageElement2.TL3_VN = null;
            isLanguageElement2.TL4_JP = null;
            isLanguageElement2.TL5_XAA = null;
            this.BUY_USER2_EMAIL.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement2});
            this.BUY_USER2_EMAIL.PromptWidth = 160;
            this.BUY_USER2_EMAIL.Size = new System.Drawing.Size(390, 21);
            this.BUY_USER2_EMAIL.TabIndex = 3;
            this.BUY_USER2_EMAIL.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.BUY_USER2_EMAIL.TextValue = "";
            this.BUY_USER2_EMAIL.Updatable = false;
            // 
            // isOraConnection1
            // 
            this.isOraConnection1.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isOraConnection1.OraConnectionInfo = oraConnectionInfo1;
            this.isOraConnection1.OraHost = "192.168.10.245";
            this.isOraConnection1.OraPassword = "infoflex";
            this.isOraConnection1.OraPort = "1521";
            this.isOraConnection1.OraServiceName = "BSKPROD";
            this.isOraConnection1.OraUserId = "APPS";
            // 
            // isMessageAdapter1
            // 
            this.isMessageAdapter1.OraConnection = this.isOraConnection1;
            // 
            // BTN_CLOSED
            // 
            this.BTN_CLOSED.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_CLOSED.ButtonText = "Closed";
            isLanguageElement6.Default = "Closed";
            isLanguageElement6.SiteName = null;
            isLanguageElement6.TL1_KR = "닫기";
            isLanguageElement6.TL2_CN = null;
            isLanguageElement6.TL3_VN = null;
            isLanguageElement6.TL4_JP = null;
            isLanguageElement6.TL5_XAA = null;
            this.BTN_CLOSED.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement6});
            this.BTN_CLOSED.Location = new System.Drawing.Point(278, 10);
            this.BTN_CLOSED.Name = "BTN_CLOSED";
            this.BTN_CLOSED.Size = new System.Drawing.Size(120, 20);
            this.BTN_CLOSED.TabIndex = 1;
            this.BTN_CLOSED.TabStop = false;
            this.BTN_CLOSED.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_CLOSED_ButtonClick);
            // 
            // isGroupBox3
            // 
            this.isGroupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.isGroupBox3.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isGroupBox3.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.isGroupBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.isGroupBox3.Controls.Add(this.BUY_USER2_EMAIL);
            this.isGroupBox3.Controls.Add(this.BUY_USER_EMAIL);
            this.isGroupBox3.Controls.Add(this.TAX_BILL_ISSUE_NO);
            this.isGroupBox3.Controls.Add(this.SELL_USER_EMAIL);
            this.isGroupBox3.Location = new System.Drawing.Point(5, 36);
            this.isGroupBox3.Name = "isGroupBox3";
            this.isGroupBox3.PromptText = "isGroupBox3";
            isLanguageElement5.Default = "isGroupBox3";
            isLanguageElement5.SiteName = null;
            isLanguageElement5.TL1_KR = null;
            isLanguageElement5.TL2_CN = null;
            isLanguageElement5.TL3_VN = null;
            isLanguageElement5.TL4_JP = null;
            isLanguageElement5.TL5_XAA = null;
            this.isGroupBox3.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement5});
            this.isGroupBox3.PromptVisible = false;
            this.isGroupBox3.Size = new System.Drawing.Size(403, 111);
            this.isGroupBox3.TabIndex = 0;
            // 
            // IDC_SET_RESEND_EMAIL
            // 
            isOraParamElement1.Direction = System.Data.ParameterDirection.Output;
            isOraParamElement1.MemberControl = null;
            isOraParamElement1.MemberValue = null;
            isOraParamElement1.OraDbTypeString = "VARCHAR2";
            isOraParamElement1.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement1.ParamName = "O_STATUS";
            isOraParamElement1.Size = 0;
            isOraParamElement1.SourceColumn = null;
            isOraParamElement2.Direction = System.Data.ParameterDirection.Output;
            isOraParamElement2.MemberControl = null;
            isOraParamElement2.MemberValue = null;
            isOraParamElement2.OraDbTypeString = "VARCHAR2";
            isOraParamElement2.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement2.ParamName = "O_MESSAGE";
            isOraParamElement2.Size = 0;
            isOraParamElement2.SourceColumn = null;
            isOraParamElement3.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement3.MemberControl = this.TAX_BILL_ISSUE_NO;
            isOraParamElement3.MemberValue = "EditValue";
            isOraParamElement3.OraDbTypeString = "VARCHAR2";
            isOraParamElement3.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement3.ParamName = "W_TAX_BILL_ISSUE_NO";
            isOraParamElement3.Size = 50;
            isOraParamElement3.SourceColumn = null;
            isOraParamElement4.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement4.MemberControl = this.SELL_USER_EMAIL;
            isOraParamElement4.MemberValue = "EditValue";
            isOraParamElement4.OraDbTypeString = "VARCHAR2";
            isOraParamElement4.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement4.ParamName = "P_SELL_USER_EMAIL";
            isOraParamElement4.Size = 0;
            isOraParamElement4.SourceColumn = null;
            isOraParamElement5.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement5.MemberControl = this.BUY_USER_EMAIL;
            isOraParamElement5.MemberValue = "EditValue";
            isOraParamElement5.OraDbTypeString = "VARCHAR2";
            isOraParamElement5.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement5.ParamName = "P_BUY_USER_EMAIL";
            isOraParamElement5.Size = 0;
            isOraParamElement5.SourceColumn = null;
            isOraParamElement6.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement6.MemberControl = this.BUY_USER2_EMAIL;
            isOraParamElement6.MemberValue = "EditValue";
            isOraParamElement6.OraDbTypeString = "VARCHAR2";
            isOraParamElement6.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement6.ParamName = "P_BUY_USER2_EMAIL";
            isOraParamElement6.Size = 0;
            isOraParamElement6.SourceColumn = null;
            isOraParamElement7.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement7.MemberControl = this.isAppInterfaceAdv1;
            isOraParamElement7.MemberValue = "SOB_ID";
            isOraParamElement7.OraDbTypeString = "NUMBER";
            isOraParamElement7.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement7.ParamName = "P_SOB_ID";
            isOraParamElement7.Size = 22;
            isOraParamElement7.SourceColumn = null;
            isOraParamElement8.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement8.MemberControl = this.isAppInterfaceAdv1;
            isOraParamElement8.MemberValue = "ORG_ID";
            isOraParamElement8.OraDbTypeString = "NUMBER";
            isOraParamElement8.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement8.ParamName = "P_ORG_ID";
            isOraParamElement8.Size = 22;
            isOraParamElement8.SourceColumn = null;
            isOraParamElement9.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement9.MemberControl = this.isAppInterfaceAdv1;
            isOraParamElement9.MemberValue = "USER_ID";
            isOraParamElement9.OraDbTypeString = "NUMBER";
            isOraParamElement9.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement9.ParamName = "P_USER_ID";
            isOraParamElement9.Size = 22;
            isOraParamElement9.SourceColumn = null;
            this.IDC_SET_RESEND_EMAIL.CommandParamElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraParamElement[] {
            isOraParamElement1,
            isOraParamElement2,
            isOraParamElement3,
            isOraParamElement4,
            isOraParamElement5,
            isOraParamElement6,
            isOraParamElement7,
            isOraParamElement8,
            isOraParamElement9});
            this.IDC_SET_RESEND_EMAIL.DataTransaction = null;
            this.IDC_SET_RESEND_EMAIL.OraConnection = this.isOraConnection1;
            this.IDC_SET_RESEND_EMAIL.OraOwner = "APPS";
            this.IDC_SET_RESEND_EMAIL.OraPackage = "TAX_BILL_ISSUE_G";
            this.IDC_SET_RESEND_EMAIL.OraProcedure = "SET_RESEND_EMAIL";
            // 
            // BTN_RESEND_EMAIL
            // 
            this.BTN_RESEND_EMAIL.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_RESEND_EMAIL.ButtonText = "Email 재전송";
            isLanguageElement1.Default = "Resend Email";
            isLanguageElement1.SiteName = null;
            isLanguageElement1.TL1_KR = "Email 재전송";
            isLanguageElement1.TL2_CN = null;
            isLanguageElement1.TL3_VN = null;
            isLanguageElement1.TL4_JP = null;
            isLanguageElement1.TL5_XAA = null;
            this.BTN_RESEND_EMAIL.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement1});
            this.BTN_RESEND_EMAIL.Location = new System.Drawing.Point(152, 10);
            this.BTN_RESEND_EMAIL.Name = "BTN_RESEND_EMAIL";
            this.BTN_RESEND_EMAIL.Size = new System.Drawing.Size(120, 20);
            this.BTN_RESEND_EMAIL.TabIndex = 0;
            this.BTN_RESEND_EMAIL.TabStop = false;
            this.BTN_RESEND_EMAIL.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.BTN_RESEND_EMAIL.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_RESEND_EMAIL_ButtonClick);
            this.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox3)).EndInit();
            this.isGroupBox3.ResumeLayout(false);

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISOraConnection isOraConnection1;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter isMessageAdapter1;
        private InfoSummit.Win.ControlAdv.ISGroupBox isGroupBox3;
        private InfoSummit.Win.ControlAdv.ISEditAdv TAX_BILL_ISSUE_NO;
        private InfoSummit.Win.ControlAdv.ISButton BTN_CLOSED;
        private InfoSummit.Win.ControlAdv.ISEditAdv SELL_USER_EMAIL;
        private InfoSummit.Win.ControlAdv.ISDataCommand IDC_SET_RESEND_EMAIL;
        private InfoSummit.Win.ControlAdv.ISEditAdv BUY_USER2_EMAIL;
        private InfoSummit.Win.ControlAdv.ISEditAdv BUY_USER_EMAIL;
        private InfoSummit.Win.ControlAdv.ISButton BTN_RESEND_EMAIL;
    }
}

