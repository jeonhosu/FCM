namespace FCMF0270
{
    partial class FCMF0270_SLIP
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
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement7 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement6 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement5 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement2 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement3 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement4 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement1 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement1 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement2 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement3 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement2 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.isMessageAdapter1 = new InfoSummit.Win.ControlAdv.ISMessageAdapter(this.components);
            this.BTN_SET_SLIP = new InfoSummit.Win.ControlAdv.ISButton();
            this.BTN_CANCEL = new InfoSummit.Win.ControlAdv.ISButton();
            this.isGroupBox2 = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.W_PERIOD_NAME_TO = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.W_SLIP_DATE = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.W_SLIP_REMARK = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.W_PERIOD_NAME_FR = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.ILA_YYYYMM_FR = new InfoSummit.Win.ControlAdv.ISLookupAdapter(this.components);
            this.ILD_YYYYMM = new InfoSummit.Win.ControlAdv.ISLookupData(this.components);
            this.ILA_YYYYMM_TO = new InfoSummit.Win.ControlAdv.ISLookupAdapter(this.components);
            this.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox2)).BeginInit();
            this.isGroupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ILA_YYYYMM_FR.PropSourceDataTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ILA_YYYYMM_TO.PropSourceDataTable)).BeginInit();
            // 
            // isAppInterfaceAdv1
            // 
            this.isAppInterfaceAdv1.AppMainButtonClick += new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv.ButtonEventHandler(this.isAppInterfaceAdv1_AppMainButtonClick);
            // 
            // isOraConnection1
            // 
            this.isOraConnection1.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isOraConnection1.OraConnectionInfo = oraConnectionInfo1;
            this.isOraConnection1.OraHost = "211.168.59.25";
            this.isOraConnection1.OraPassword = "infoflex";
            this.isOraConnection1.OraPort = "1521";
            this.isOraConnection1.OraServiceName = "FXCDB";
            this.isOraConnection1.OraUserId = "APPS";
            // 
            // isMessageAdapter1
            // 
            this.isMessageAdapter1.OraConnection = this.isOraConnection1;
            // 
            // BTN_SET_SLIP
            // 
            this.BTN_SET_SLIP.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_SET_SLIP.ButtonText = "Create Slip";
            isLanguageElement1.Default = "Create Slip";
            isLanguageElement1.SiteName = null;
            isLanguageElement1.TL1_KR = "전표 생성";
            isLanguageElement1.TL2_CN = "";
            isLanguageElement1.TL3_VN = "";
            isLanguageElement1.TL4_JP = "";
            isLanguageElement1.TL5_XAA = "";
            this.BTN_SET_SLIP.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement1});
            this.BTN_SET_SLIP.ForeColor = System.Drawing.Color.Blue;
            // 
            // FCMF0270_SLIP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(463, 148);
            this.ControlBox = false;
            this.Controls.Add(this.isGroupBox2);
            this.Controls.Add(this.BTN_CANCEL);
            this.Controls.Add(this.BTN_SET_SLIP);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FCMF0270_SLIP";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Set Slip";
            this.Load += new System.EventHandler(this.FCMF0270_SLIP_Load);
            this.Shown += new System.EventHandler(this.FCMF0270_SLIP_Shown);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FCMF0270_SLIP_FormClosed);
            this.BTN_SET_SLIP.Location = new System.Drawing.Point(191, 112);
            this.BTN_SET_SLIP.Name = "BTN_SET_SLIP";
            this.BTN_SET_SLIP.Size = new System.Drawing.Size(90, 25);
            this.BTN_SET_SLIP.TabIndex = 2;
            this.BTN_SET_SLIP.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.ibtnSAVE_ButtonClick);
            // 
            // BTN_CANCEL
            // 
            this.BTN_CANCEL.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_CANCEL.ButtonText = "Cancel";
            isLanguageElement7.Default = "Cancel";
            isLanguageElement7.SiteName = null;
            isLanguageElement7.TL1_KR = "취소";
            isLanguageElement7.TL2_CN = "";
            isLanguageElement7.TL3_VN = "";
            isLanguageElement7.TL4_JP = "";
            isLanguageElement7.TL5_XAA = "";
            this.BTN_CANCEL.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement7});
            this.BTN_CANCEL.ForeColor = System.Drawing.Color.Blue;
            this.BTN_CANCEL.Location = new System.Drawing.Point(287, 112);
            this.BTN_CANCEL.Name = "BTN_CANCEL";
            this.BTN_CANCEL.Size = new System.Drawing.Size(90, 25);
            this.BTN_CANCEL.TabIndex = 3;
            this.BTN_CANCEL.Visible = false;
            this.BTN_CANCEL.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.ibtnCANCEL_ButtonClick);
            // 
            // isGroupBox2
            // 
            this.isGroupBox2.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isGroupBox2.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.isGroupBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.isGroupBox2.Controls.Add(this.W_SLIP_DATE);
            this.isGroupBox2.Controls.Add(this.W_SLIP_REMARK);
            this.isGroupBox2.Controls.Add(this.W_PERIOD_NAME_FR);
            this.isGroupBox2.Controls.Add(this.W_PERIOD_NAME_TO);
            this.isGroupBox2.Location = new System.Drawing.Point(8, 8);
            this.isGroupBox2.Name = "isGroupBox2";
            this.isGroupBox2.PromptText = "isGroupBox1";
            isLanguageElement6.Default = "isGroupBox1";
            isLanguageElement6.SiteName = null;
            isLanguageElement6.TL1_KR = null;
            isLanguageElement6.TL2_CN = null;
            isLanguageElement6.TL3_VN = null;
            isLanguageElement6.TL4_JP = null;
            isLanguageElement6.TL5_XAA = null;
            this.isGroupBox2.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement6});
            this.isGroupBox2.PromptVisible = false;
            this.isGroupBox2.Size = new System.Drawing.Size(447, 95);
            this.isGroupBox2.TabIndex = 5;
            // 
            // W_PERIOD_NAME_TO
            // 
            this.W_PERIOD_NAME_TO.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.W_PERIOD_NAME_TO.ComboBoxValue = "";
            this.W_PERIOD_NAME_TO.ComboData = null;
            this.W_PERIOD_NAME_TO.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_PERIOD_NAME_TO.DataAdapter = null;
            this.W_PERIOD_NAME_TO.DataColumn = null;
            this.W_PERIOD_NAME_TO.DateTimeValue = new System.DateTime(2015, 5, 26, 0, 0, 0, 0);
            this.W_PERIOD_NAME_TO.DoubleValue = 0;
            this.W_PERIOD_NAME_TO.EditValue = "";
            this.W_PERIOD_NAME_TO.Location = new System.Drawing.Point(259, 10);
            this.W_PERIOD_NAME_TO.LookupAdapter = this.ILA_YYYYMM_TO;
            this.W_PERIOD_NAME_TO.Name = "W_PERIOD_NAME_TO";
            this.W_PERIOD_NAME_TO.Nullable = true;
            this.W_PERIOD_NAME_TO.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_PERIOD_NAME_TO.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_PERIOD_NAME_TO.PromptText = "~";
            isLanguageElement5.Default = "~";
            isLanguageElement5.SiteName = null;
            isLanguageElement5.TL1_KR = "~";
            isLanguageElement5.TL2_CN = "";
            isLanguageElement5.TL3_VN = "";
            isLanguageElement5.TL4_JP = "";
            isLanguageElement5.TL5_XAA = "";
            this.W_PERIOD_NAME_TO.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement5});
            this.W_PERIOD_NAME_TO.PromptWidth = 20;
            this.W_PERIOD_NAME_TO.Size = new System.Drawing.Size(170, 21);
            this.W_PERIOD_NAME_TO.TabIndex = 3;
            this.W_PERIOD_NAME_TO.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.W_PERIOD_NAME_TO.TextValue = "";
            // 
            // W_SLIP_DATE
            // 
            this.W_SLIP_DATE.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.W_SLIP_DATE.ComboBoxValue = "";
            this.W_SLIP_DATE.ComboData = null;
            this.W_SLIP_DATE.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_SLIP_DATE.DataAdapter = null;
            this.W_SLIP_DATE.DataColumn = null;
            this.W_SLIP_DATE.DateTimeValue = new System.DateTime(2015, 5, 26, 0, 0, 0, 0);
            this.W_SLIP_DATE.DoubleValue = 0;
            this.W_SLIP_DATE.EditAdvType = InfoSummit.Win.ControlAdv.ISUtil.Enum.EditAdvType.DateTimeEdit;
            this.W_SLIP_DATE.EditValue = null;
            this.W_SLIP_DATE.Location = new System.Drawing.Point(3, 65);
            this.W_SLIP_DATE.LookupAdapter = null;
            this.W_SLIP_DATE.Name = "W_SLIP_DATE";
            this.W_SLIP_DATE.Nullable = true;
            this.W_SLIP_DATE.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_SLIP_DATE.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_SLIP_DATE.PromptText = "회계일자";
            isLanguageElement2.Default = "Slip Date";
            isLanguageElement2.SiteName = null;
            isLanguageElement2.TL1_KR = "회계일자";
            isLanguageElement2.TL2_CN = "";
            isLanguageElement2.TL3_VN = "";
            isLanguageElement2.TL4_JP = "";
            isLanguageElement2.TL5_XAA = "";
            this.W_SLIP_DATE.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement2});
            this.W_SLIP_DATE.Size = new System.Drawing.Size(250, 21);
            this.W_SLIP_DATE.TabIndex = 5;
            this.W_SLIP_DATE.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.W_SLIP_DATE.TextValue = "";
            // 
            // W_SLIP_REMARK
            // 
            this.W_SLIP_REMARK.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.W_SLIP_REMARK.ComboBoxValue = "";
            this.W_SLIP_REMARK.ComboData = null;
            this.W_SLIP_REMARK.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_SLIP_REMARK.DataAdapter = null;
            this.W_SLIP_REMARK.DataColumn = "";
            this.W_SLIP_REMARK.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.W_SLIP_REMARK.DoubleValue = 0;
            this.W_SLIP_REMARK.EditValue = "기간비용 대체";
            this.W_SLIP_REMARK.Location = new System.Drawing.Point(3, 37);
            this.W_SLIP_REMARK.LookupAdapter = null;
            this.W_SLIP_REMARK.Name = "W_SLIP_REMARK";
            this.W_SLIP_REMARK.Nullable = true;
            this.W_SLIP_REMARK.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_SLIP_REMARK.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_SLIP_REMARK.PromptText = "적요";
            isLanguageElement3.Default = "Remark";
            isLanguageElement3.SiteName = null;
            isLanguageElement3.TL1_KR = "적요";
            isLanguageElement3.TL2_CN = "";
            isLanguageElement3.TL3_VN = "";
            isLanguageElement3.TL4_JP = "";
            isLanguageElement3.TL5_XAA = "";
            this.W_SLIP_REMARK.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement3});
            this.W_SLIP_REMARK.Size = new System.Drawing.Size(426, 21);
            this.W_SLIP_REMARK.TabIndex = 4;
            this.W_SLIP_REMARK.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.W_SLIP_REMARK.TextValue = "기간비용 대체";
            // 
            // W_PERIOD_NAME_FR
            // 
            this.W_PERIOD_NAME_FR.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.W_PERIOD_NAME_FR.ComboBoxValue = "";
            this.W_PERIOD_NAME_FR.ComboData = null;
            this.W_PERIOD_NAME_FR.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_PERIOD_NAME_FR.DataAdapter = null;
            this.W_PERIOD_NAME_FR.DataColumn = null;
            this.W_PERIOD_NAME_FR.DateTimeValue = new System.DateTime(2015, 5, 26, 0, 0, 0, 0);
            this.W_PERIOD_NAME_FR.DoubleValue = 0;
            this.W_PERIOD_NAME_FR.EditValue = "";
            this.W_PERIOD_NAME_FR.Location = new System.Drawing.Point(3, 10);
            this.W_PERIOD_NAME_FR.LookupAdapter = this.ILA_YYYYMM_FR;
            this.W_PERIOD_NAME_FR.Name = "W_PERIOD_NAME_FR";
            this.W_PERIOD_NAME_FR.Nullable = true;
            this.W_PERIOD_NAME_FR.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_PERIOD_NAME_FR.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.W_PERIOD_NAME_FR.PromptText = "처리기간";
            isLanguageElement4.Default = "Period Date";
            isLanguageElement4.SiteName = null;
            isLanguageElement4.TL1_KR = "처리기간";
            isLanguageElement4.TL2_CN = "";
            isLanguageElement4.TL3_VN = "";
            isLanguageElement4.TL4_JP = "";
            isLanguageElement4.TL5_XAA = "";
            this.W_PERIOD_NAME_FR.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement4});
            this.W_PERIOD_NAME_FR.Size = new System.Drawing.Size(250, 21);
            this.W_PERIOD_NAME_FR.TabIndex = 1;
            this.W_PERIOD_NAME_FR.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.W_PERIOD_NAME_FR.TextValue = "";
            // 
            // ILA_YYYYMM_FR
            // 
            this.ILA_YYYYMM_FR.DisplayCaption = "처리시작월";
            isOraColElement1.DataColumn = "YYYYMM";
            isOraColElement1.DataOrdinal = 0;
            isOraColElement1.DataType = "System.String";
            isOraColElement1.HeaderPrompt = "Period Name(Fr)";
            isOraColElement1.LastValue = null;
            isOraColElement1.MemberControl = this.W_PERIOD_NAME_FR;
            isOraColElement1.MemberValue = "EditValue";
            isOraColElement1.Nullable = null;
            isOraColElement1.Ordinal = 0;
            isOraColElement1.RelationKeyColumn = null;
            isOraColElement1.ReturnParameter = null;
            isOraColElement1.TL1_KR = "처리년월(시작)";
            isOraColElement1.TL2_CN = null;
            isOraColElement1.TL3_VN = null;
            isOraColElement1.TL4_JP = null;
            isOraColElement1.TL5_XAA = null;
            isOraColElement1.Visible = 1;
            isOraColElement1.Width = 120;
            this.ILA_YYYYMM_FR.LookupColElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraColElement[] {
            isOraColElement1});
            this.ILA_YYYYMM_FR.LookupData = this.ILD_YYYYMM;
            this.ILA_YYYYMM_FR.LookupSize = new System.Drawing.Size(186, 344);
            this.ILA_YYYYMM_FR.SelectLookupSize = InfoSummit.Win.ControlAdv.ISUtil.Enum.SelectLookupSize.Custom;
            // 
            // ILD_YYYYMM
            // 
            isOraParamElement1.Direction = System.Data.ParameterDirection.Output;
            isOraParamElement1.MemberControl = null;
            isOraParamElement1.MemberValue = null;
            isOraParamElement1.OraDbTypeString = "REF CURSOR";
            isOraParamElement1.OraType = System.Data.OracleClient.OracleType.Cursor;
            isOraParamElement1.ParamName = "P_CURSOR";
            isOraParamElement1.Size = 0;
            isOraParamElement1.SourceColumn = null;
            isOraParamElement2.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement2.MemberControl = null;
            isOraParamElement2.MemberValue = null;
            isOraParamElement2.OraDbTypeString = "VARCHAR2";
            isOraParamElement2.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement2.ParamName = "W_START_YYYYMM";
            isOraParamElement2.Size = 7;
            isOraParamElement2.SourceColumn = null;
            isOraParamElement3.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement3.MemberControl = null;
            isOraParamElement3.MemberValue = null;
            isOraParamElement3.OraDbTypeString = "VARCHAR2";
            isOraParamElement3.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement3.ParamName = "W_END_YYYYMM";
            isOraParamElement3.Size = 7;
            isOraParamElement3.SourceColumn = null;
            this.ILD_YYYYMM.LookupParamElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraParamElement[] {
            isOraParamElement1,
            isOraParamElement2,
            isOraParamElement3});
            this.ILD_YYYYMM.OraConnection = this.isOraConnection1;
            this.ILD_YYYYMM.OraOwner = "APPS";
            this.ILD_YYYYMM.OraPackage = "EAPP_CALENDAR_G";
            this.ILD_YYYYMM.OraProcedure = "LU_CALENDAR_YYYYMM";
            // 
            // ILA_YYYYMM_TO
            // 
            this.ILA_YYYYMM_TO.DisplayCaption = "처리종료월";
            isOraColElement2.DataColumn = "YYYYMM";
            isOraColElement2.DataOrdinal = 0;
            isOraColElement2.DataType = "System.String";
            isOraColElement2.HeaderPrompt = "Period Name(To)";
            isOraColElement2.LastValue = null;
            isOraColElement2.MemberControl = this.W_PERIOD_NAME_TO;
            isOraColElement2.MemberValue = "EditValue";
            isOraColElement2.Nullable = null;
            isOraColElement2.Ordinal = 0;
            isOraColElement2.RelationKeyColumn = null;
            isOraColElement2.ReturnParameter = null;
            isOraColElement2.TL1_KR = "처리년월(종료)";
            isOraColElement2.TL2_CN = null;
            isOraColElement2.TL3_VN = null;
            isOraColElement2.TL4_JP = null;
            isOraColElement2.TL5_XAA = null;
            isOraColElement2.Visible = 1;
            isOraColElement2.Width = 120;
            this.ILA_YYYYMM_TO.LookupColElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraColElement[] {
            isOraColElement2});
            this.ILA_YYYYMM_TO.LookupData = this.ILD_YYYYMM;
            this.ILA_YYYYMM_TO.LookupSize = new System.Drawing.Size(186, 344);
            this.ILA_YYYYMM_TO.SelectLookupSize = InfoSummit.Win.ControlAdv.ISUtil.Enum.SelectLookupSize.Custom;
            this.ILA_YYYYMM_TO.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox2)).EndInit();
            this.isGroupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ILA_YYYYMM_FR.PropSourceDataTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ILA_YYYYMM_TO.PropSourceDataTable)).EndInit();

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISOraConnection isOraConnection1;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter isMessageAdapter1;
        private InfoSummit.Win.ControlAdv.ISButton BTN_SET_SLIP;
        private InfoSummit.Win.ControlAdv.ISButton BTN_CANCEL;
        private InfoSummit.Win.ControlAdv.ISGroupBox isGroupBox2;
        private InfoSummit.Win.ControlAdv.ISEditAdv W_PERIOD_NAME_TO;
        private InfoSummit.Win.ControlAdv.ISEditAdv W_SLIP_DATE;
        private InfoSummit.Win.ControlAdv.ISEditAdv W_SLIP_REMARK;
        private InfoSummit.Win.ControlAdv.ISEditAdv W_PERIOD_NAME_FR;
        private InfoSummit.Win.ControlAdv.ISLookupAdapter ILA_YYYYMM_FR;
        private InfoSummit.Win.ControlAdv.ISLookupData ILD_YYYYMM;
        private InfoSummit.Win.ControlAdv.ISLookupAdapter ILA_YYYYMM_TO;
    }
}

