using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
using System.IO;
using System.Net;
using System.Xml;

namespace FCMF0131
{
    public partial class FCMF0131 : Office2007Form
    {
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        private ISFileTransferAdv mFileTransfer;
        private isFTP_Info mFTP_Info;

        string mSOURCE_CATEGORY = "CUST_ATTACH_FILE";
        string mFTP_INFO_CODE = "CUST_ATTACH_FILE";           //거래처등록 첨부파일//

        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 실행 디렉토리.        
        private string mDownload_Folder = string.Empty;                                   // 거래처등록 다운로드 폴더
        private bool mFTP_Connect_Status = false;                                           // FTP 정보 상태
         
        #region ----- Constructor -----

        public FCMF0131(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void FCMF0131_Load(object sender, EventArgs e)
        {
            V_VENDOR_TYPE.EditValue = Convert.ToString("A");
            V_RD_ALL.Checked = true;

            Set_FTP_Info();
            V_ATTACH_SOURCE_CATEGORY.EditValue = "CUST_ATTACH_FILE";

            IDA_VENDOR.FillSchema();
            idaCUST_SHIP_TO.FillSchema();
            idaCUST_PERSON.FillSchema();
        }

        //private void Insert_Bank_Account()
        //{
        //    isgBANK_ACCOUNT.SetCellValue("ENABLED_FLAG", "Y");
        //    isgBANK_ACCOUNT.SetCellValue("EFFECTIVE_DATE_FR", DateTime.Today);
        //}

        private void Show_Address()
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();
            
            DialogResult dlgRESULT;

            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, H_ZIP_CODE.EditValue, H_ADDRESS1.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vEAPF0299, isAppInterfaceAdv1.AppInterface);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                H_ZIP_CODE.EditValue = vEAPF0299.Get_Zip_Code;
                H_ADDRESS1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void Show_Tax_Reg_Num(bool pVisible)
        { 
            BTN_TAX_REG_NUM_CHK.Visible = pVisible;
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    AR_CUSTOMER_INQUIRY();
                    // Select Error Routine
                    if (IDA_VENDOR.GetSelectParamValue("X_ERR_MSG") != null)
                    {
                        MessageBoxAdv.Show(IDA_VENDOR.GetSelectParamValue("X_ERR_MSG").ToString());
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_VENDOR.IsFocused)    // AR_CUSTOMER
                    {
                        IDA_VENDOR.AddOver();
                        Form_DefaultValue();
                    }
                    else if (idaCUST_SHIP_TO.IsFocused && IDA_VENDOR.CurrentRow != null)    // AR_CUSTOMER_SHIP_TO
                    {
                        idaCUST_SHIP_TO.AddOver();
                        isgCUST_SHIP_TO.SetCellValue("EFFECTIVE_DATE_FR", DateTime.Today);
                        isgCUST_SHIP_TO.SetCellValue("ENABLED_FLAG", "Y");
                    }
                    else if (idaCUST_PERSON.IsFocused && IDA_VENDOR.CurrentRow != null)    // AR_CUSTOMER_PERSON
                    {
                        idaCUST_PERSON.AddOver();
                        isgCUST_PERSON.SetCellValue("EFFECTIVE_DATE_FR", DateTime.Today);
                        isgCUST_PERSON.SetCellValue("ENABLED_FLAG", "Y");
                    }
                    //if (idaCUST_BANK_ACCT.IsFocused && IDA_VENDOR.CurrentRow != null)
                    //{
                    //    idaCUST_BANK_ACCT.AddOver();
                    //    isgBANK_ACCOUNT.SetCellValue("EFFECTIVE_DATE_FR", DateTime.Today);
                    //    isgBANK_ACCOUNT.SetCellValue("ENABLED_FLAG", "Y");
                    //    Insert_Bank_Account();
                    //}
                    else if (IDA_VENDOR_HISTORY.IsFocused)
                    {
                        IDA_VENDOR_HISTORY.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_VENDOR.IsFocused)    // AR_CUSTOMER
                    {
                        IDA_VENDOR.AddUnder();
                        Form_DefaultValue();
                    }
                    else if (idaCUST_SHIP_TO.IsFocused && IDA_VENDOR.CurrentRow != null)    // AR_CUSTOMER_SHIP_TO
                    {
                        idaCUST_SHIP_TO.AddUnder();
                        isgCUST_SHIP_TO.SetCellValue("EFFECTIVE_DATE_FR", DateTime.Today);
                        isgCUST_SHIP_TO.SetCellValue("ENABLED_FLAG", "Y");
                    }
                    else if (idaCUST_PERSON.IsFocused && IDA_VENDOR.CurrentRow != null)    // AR_CUSTOMER_PERSON
                    {
                        idaCUST_PERSON.AddUnder();
                        isgCUST_PERSON.SetCellValue("EFFECTIVE_DATE_FR", DateTime.Today);
                        isgCUST_PERSON.SetCellValue("ENABLED_FLAG", "Y");
                    }
                    //else if (idaCUST_BANK_ACCT.IsFocused && IDA_VENDOR.CurrentRow != null)
                    //{
                    //    idaCUST_BANK_ACCT.AddUnder();
                    //    isgBANK_ACCOUNT.SetCellValue("EFFECTIVE_DATE_FR", DateTime.Today);
                    //    isgBANK_ACCOUNT.SetCellValue("ENABLED_FLAG", "Y");
                    //    Insert_Bank_Account();
                    //} 
                    else if (IDA_VENDOR_HISTORY.IsFocused)
                    {
                        IDA_VENDOR_HISTORY.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    //if (idaCUSTOMER_SITE.Refillable == false)
                    //{
                        //if (IDA_VENDOR.ChangedRowCount > 0 || IDA_VENDOR_HISTORY.ChangedRowCount > 0 || idaBANK_ACCOUNT.ChangedRowCount > 0 || idaCUST_PERSON.ChangedRowCount > 0 || idaCUST_SHIP_TO.ChangedRowCount > 0)
                        //{
                            DialogResult vResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10150"), "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                            if (vResult == DialogResult.OK)
                            {
                                IDA_VENDOR.Update();

                                if (Pre_Update() == false)
                                {
                                    return;
                                } 

                                //if (Convert.ToString(IDA_VENDOR.GetInsertParamValue("O_CUST_CODE")) != "")
                                //{
                                //    H_VENDOR_CODE.EditValue = IDA_VENDOR.GetInsertParamValue("O_CUST_CODE");
                                //    IDA_VENDOR.CurrentRow.AcceptChanges();
                                //    IDA_VENDOR.Refillable = true;
                                //}

                                if (IDA_VENDOR_HISTORY.IsFocused)
                                {
                                    IDA_VENDOR_HISTORY.Update();
                                }
                            }
                            else if (vResult == DialogResult.Cancel)
                            {
                                
                            }
                        //}

                        
                        
                        // Insert Error Routine
                        //if (idaCUSTOMER_SITE.GetInsertParamValue("X_ERR_MSG") != null)
                        //{
                        //    MessageBoxAdv.Show(idaCUSTOMER_SITE.GetInsertParamValue("X_ERR_MSG").ToString());
                        //}
                        //// Update Error Routine
                        //if (idaCUSTOMER_SITE.GetUpdateParamValue("X_ERR_MSG") != null)
                        //{
                        //    MessageBoxAdv.Show(idaCUSTOMER_SITE.GetUpdateParamValue("X_ERR_MSG").ToString());
                        //}
                    //}

                    //if (idaCUST_SHIP_TO.Refillable == false)
                    //{
                    //    idaCUST_SHIP_TO.Update();
                    //    // Insert Error Routine
                    //    if (idaCUST_SHIP_TO.GetInsertParamValue("X_ERR_MSG") != null)
                    //    {
                    //        MessageBoxAdv.Show(idaCUST_SHIP_TO.GetInsertParamValue("X_ERR_MSG").ToString());
                    //    }
                    //    // Update Error Routine
                    //    if (idaCUST_SHIP_TO.GetUpdateParamValue("X_ERR_MSG") != null)
                    //    {
                    //        MessageBoxAdv.Show(idaCUST_SHIP_TO.GetUpdateParamValue("X_ERR_MSG").ToString());
                    //    }
                    //}

                    //if (idaCUST_PERSON.Refillable == false)
                    //{
                    //    idaCUST_PERSON.Update();
                    //    // Insert Error Routine
                    //    if (idaCUST_PERSON.GetInsertParamValue("X_ERR_MSG") != null)
                    //    {
                    //        MessageBoxAdv.Show(idaCUST_PERSON.GetInsertParamValue("X_ERR_MSG").ToString());
                    //    }
                    //    // Update Error Routine
                    //    if (idaCUST_PERSON.GetUpdateParamValue("X_ERR_MSG") != null)
                    //    {
                    //        MessageBoxAdv.Show(idaCUST_PERSON.GetUpdateParamValue("X_ERR_MSG").ToString());
                    //    }
                    //}
                    //if (idaCUST_BANK_ACCT.Refillable == false)
                    //{
                    //    idaCUST_BANK_ACCT.Update();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_VENDOR.IsFocused)
                    {
                        IDA_VENDOR.Cancel();
                    }
                    else if (idaCUST_SHIP_TO.IsFocused)
                    {
                        idaCUST_SHIP_TO.Cancel();
                    }
                    else if (idaCUST_PERSON.IsFocused)
                    {
                        idaCUST_PERSON.Cancel();
                    }
                    else if (IDA_VENDOR_HISTORY.IsFocused)
                    {
                        IDA_VENDOR_HISTORY.Cancel();
                    }
                    //else if (idaBANK_ACCOUNT.IsFocused)
                    //{
                    //    idaBANK_ACCOUNT.Cancel();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    //if (idaCUST_BANK_ACCT.IsFocused)
                    //{
                    //    idaCUST_BANK_ACCT.Delete();
                    //}
                    if (IDA_VENDOR.IsFocused)
                    {
                        if (IDA_VENDOR.CurrentRow.RowState == DataRowState.Added)
                        {
                            IDA_VENDOR.Delete();
                        }
                    }
                    if (IDA_VENDOR_HISTORY.IsFocused)
                    {
                        IDA_VENDOR_HISTORY.Delete();
                    }
                }
            }
        }

        #endregion;

        #region -- Data Find --

        private void AR_CUSTOMER_INQUIRY()
        {
            IDA_VENDOR.Fill();

            H_VENDOR_CODE.Focus();
        }

        private void SEARCH_DB_ATTACHMENT(object pSOURCE_ID)
        {
            //이미지 초기화;
            //ImageView(string.Empty);

            //첨부파일 리스트 조회 
            IDA_ATTACH_FILE.SetSelectParamValue("P_SOURCE_CATEGORY", "CUST_ATTACH_FILE");
            IDA_ATTACH_FILE.SetSelectParamValue("P_SOURCE_ID", pSOURCE_ID);
            IDA_ATTACH_FILE.Refillable = true;
            IDA_ATTACH_FILE.Fill();
        }

        #endregion

        #region -- Default Value Setting --

        private void Form_DefaultValue()
        {
            idcLOCAL_DATE.ExecuteNonQuery();
            H_EFFECTIVE_DATE_FR.EditValue = idcLOCAL_DATE.GetCommandParamValue("X_LOCAL_DATE");
            T_ENABLED_FLAG.CheckBoxValue = "Y";

            H_VENDOR_CODE.Focus();
        }

        #endregion

        #region ---- pre update ---- 

        private bool Pre_Update()
        {
            bool CHK_RESULT = true;

            //if (idaCUSTOMER_SITE.Refillable == false)
            //{
            //    //고객코드
            //    if (string.IsNullOrEmpty(CUST_SITE_CODE.EditValue.ToString()))
            //    {
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Customer Code"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        CHK_RESULT = false;
            //        return CHK_RESULT;
            //    }

            //    //고객명
            //    if (string.IsNullOrEmpty(CUST_SITE_FULL_NAME.EditValue.ToString()))
            //    {
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Customer Full Name"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        CHK_RESULT = false;
            //        return CHK_RESULT;
            //    }

            //    //약명
            //    if (string.IsNullOrEmpty(CUST_SITE_SHORT_NAME.EditValue.ToString()))
            //    {
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Customer Short Name"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        CHK_RESULT = false;
            //        return CHK_RESULT;
            //    }

            //    //고객 Party
            //    if (string.IsNullOrEmpty(CUST_PARTY_DESC.EditValue.ToString()))
            //    {
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Customer Party"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        CHK_RESULT = false;
            //        return CHK_RESULT;
            //    }

            //    //국가
            //    if (string.IsNullOrEmpty(COUNTRY_CODE.EditValue.ToString()))
            //    {
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Country"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        CHK_RESULT = false;
            //        return CHK_RESULT;
            //    }

            //    //거래시작일
            //    if (EFFECTIVE_DATE_FR.EditValue == null)
            //    {
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Effective Date(F)"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        CHK_RESULT = false;
            //        return CHK_RESULT;
            //    }
            //}

            if (idaCUST_SHIP_TO.Refillable == false)
            {
                
            }

            if (idaCUST_PERSON.Refillable == false)
            {

            }
            return CHK_RESULT;
        }

        #endregion


        #region ----- FTP Infomation -----
        //ftp 접속정보 및 환경 정보 설정 

        //첨부파일
        private void Set_FTP_Info()
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            mFTP_Connect_Status = false;
            try
            {
                IDC_FTP_INFO.SetCommandParamValue("W_FTP_INFO_CODE", "CUST_ATTACH_FILE");
                IDC_FTP_INFO.ExecuteNonQuery();
                if (IDC_FTP_INFO.ExcuteError)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();
                    return;
                }

                mFTP_Info = new isFTP_Info();

                mFTP_Info.Host = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_FTP_IP"));
                mFTP_Info.Port = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_FTP_PORT"));
                mFTP_Info.UserID = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_FTP_USER_ID"));
                mFTP_Info.Password = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_FTP_PASSWORD"));
                mFTP_Info.Passive = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_FTP_USEPASSIVE"));
                mFTP_Info.FTP_Folder = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_FTP_SOURCEPATH"));
                mFTP_Info.Client_Folder = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_CLIENT_TARGETPATH"));
            }
            catch (Exception Ex)
            {
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            if (mFTP_Info.Host == string.Empty)
            {
                //ftp접속정보 오류          
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            try
            {
                //FileTransfer Initialze
                mFileTransfer = new ISFileTransferAdv();
                mFileTransfer.Host = mFTP_Info.Host;
                mFileTransfer.Port = mFTP_Info.Port;
                mFileTransfer.UserId = mFTP_Info.UserID;
                mFileTransfer.Password = mFTP_Info.Password;
                if (mFTP_Info.Passive == "Y")
                {
                    mFileTransfer.UsePassive = true;
                }
                else
                {
                    mFileTransfer.UsePassive = false;
                }

                mDownload_Folder = string.Format("{0}\\{1}", mClient_Base_Path, mFTP_Info.Client_Folder);
            }
            catch (System.Exception Ex)
            {
                //ftp접속정보 오류 
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            //Client Download Folder 없으면 생성 
            System.IO.DirectoryInfo vDownload_Folder = new System.IO.DirectoryInfo(mDownload_Folder);
            if (vDownload_Folder.Exists == false) //있으면 True, 없으면 False
            {
                vDownload_Folder.Create();
            }

            mFTP_Connect_Status = true;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion

        #region ----- File Upload Methods -----
        //업로드 메소드
        private bool UpLoadFile(object pVENDOR_ID, object pVENDOR_CODE, string pPlus_Name)
        {
            bool isUpload = false;

            if (mFTP_Connect_Status == false)
            {
                isAppInterfaceAdv1.OnAppMessage("FTP Server Connect Fail. Check FTP Server");
                return isUpload;
            }

            if (iConv.ISNull(pVENDOR_CODE) != string.Empty)
            {
                string vSTATUS = "F";
                string vMESSAGE = string.Empty;


                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "All File(*.*) | *.*| Excel File(*.xls; *.xlsx)| *.xls; *.xlsx | PowerPoint File(*.ppt; *.pptx)| *.ppt; *.pptx | jpg file(*.jpg) | *.jpg";
                openFileDialog1.DefaultExt = "*.*";
                openFileDialog1.FileName = "";
                openFileDialog1.Multiselect = false;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        //FileTransfer Initialze
                        mFileTransfer = new ISFileTransferAdv();
                        mFileTransfer.Host = mFTP_Info.Host;
                        mFileTransfer.Port = mFTP_Info.Port;
                        mFileTransfer.UserId = mFTP_Info.UserID;
                        mFileTransfer.Password = mFTP_Info.Password; 
                        if (mFTP_Info.Passive == "Y")
                        {
                            mFileTransfer.UsePassive = true;
                        }
                        else
                        {
                            mFileTransfer.UsePassive = false;
                        }
                    }
                    catch (System.Exception Ex)
                    {
                        //ftp접속정보 오류 
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                        Application.UseWaitCursor = false;
                        this.Cursor = Cursors.Default;
                        Application.DoEvents();
                        return isUpload;
                    }

                    //1. 사용자 선택 파일 
                    string vSelectFullPath = openFileDialog1.FileName;
                    string vSelectDirectoryPath = Path.GetDirectoryName(openFileDialog1.FileName);

                    string vFileName = Path.GetFileName(openFileDialog1.FileName);
                    string vFileExtension = Path.GetExtension(openFileDialog1.FileName).ToUpper();

                    //if (vFileExtension != ".JPG" && vFileExtension != ".BMP" && vFileExtension != ".GIF")
                    //{
                    //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    //    return false;
                    //}

                    //2. 첨부파일 DB저장
                    //IDC_INSERT_DOC_ATTACHMENT_IMAGE.DataTransaction = isDataTransaction1;
                    //IDC_INSERT_DOC_ATTACHMENT_IMAGE_LOG.DataTransaction = isDataTransaction1; 

                    ////FILE VERSION 
                    //IDC_GET_ATTACHMENT_VERSION.SetCommandParamValue("P_SOURCE_CATEGORY", "SDM_IMAGE");
                    //IDC_GET_ATTACHMENT_VERSION.SetCommandParamValue("P_SOURCE_ID", pINVENTORY_ITEM_ID);
                    ////IDC_INSERT_DOC_ATTACHMENT_IMAGE.SetCommandParamValue("P_DOC_CATEGORY_LCODE", pDOC_CATEGORY_LCODE);
                    ////IDC_INSERT_DOC_ATTACHMENT_IMAGE.SetCommandParamValue("P_DOC_TYPE_LCODE", pDOC_TYPE_LCODE);
                    //IDC_GET_ATTACHMENT_VERSION.ExecuteNonQuery();
                    //object vVERSION_SEQ = IDC_GET_ATTACHMENT_VERSION.GetCommandParamValue("O_VERSION_SEQ");

                    //FTP FILE NAME 
                    object vFTP_FILE_NAME = string.Format("{0}{1}", pVENDOR_CODE, pPlus_Name);

                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_CATEGORY", "CUST_ATTACH_FILE"); //구분
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_ID", pVENDOR_ID); //구분
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_USER_FILE_NAME", vFileName);
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_FTP_FILE_NAME", vFTP_FILE_NAME);
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_EXTENSION_NAME", vFileExtension);
                    //IDC_INSERT_DOC_ATTACHMENT_IMAGE.SetCommandParamValue("P_DOC_CATEGORY_LCODE", pDOC_CATEGORY_LCODE);
                    //IDC_INSERT_DOC_ATTACHMENT_IMAGE.SetCommandParamValue("P_DOC_TYPE_LCODE", pDOC_TYPE_LCODE);
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_IN_OUT_STATUS", "IN");
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_APPROVAL_REQ_FLAG", "N");
                    IDC_INSERT_DOC_ATTACHMENT.ExecuteNonQuery();

                    vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_MESSAGE"));
                    object vDOC_ATTACHMENT_ID = IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_DOC_ATTACHMENT_ID");
                    vFTP_FILE_NAME = IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_FTP_FILE_NAME");


                    if (IDC_INSERT_DOC_ATTACHMENT.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = Cursors.Default;
                        Application.DoEvents();
                         
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return isUpload;
                    }

                    ////3. 첨부파일 로그 저장
                    IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", vDOC_ATTACHMENT_ID);
                    IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "IN");
                    IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = Cursors.Default;
                        Application.DoEvents();
                         
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return isUpload;
                    }

                    //4. 파일 업로드
                    try
                    {
                        int vArryCount = openFileDialog1.FileNames.Length;
                        for (int r = 0; r < vArryCount; r++)
                        {
                            mFileTransfer.ShowProgress = true;      //진행바 보이기 

                            //업로드 환경 설정 
                            mFileTransfer.SourceDirectory = vSelectDirectoryPath;
                            mFileTransfer.SourceFileName = vFileName;
                            mFileTransfer.TargetDirectory = mFTP_Info.FTP_Folder;
                            mFileTransfer.TargetFileName = iConv.ISNull(vFTP_FILE_NAME);

                            bool isUpLoad = mFileTransfer.Upload();

                            if (isUpLoad == true)
                            {
                                isUpload = true;
                            }
                            else
                            {
                                isUpload = false; 
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return isUpload;
                            }
                        } 
                    }

                    catch (Exception Ex)
                    { 
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                        return isUpload;
                    }
                }
            }
            return isUpload;
        }

        #endregion

        #region ----- is View file Method -----

        private string isDownload(object pFILE_ENTRY_ID, string pSAVE_FileName, string pFTP_FILE_NAME, string vSAVE_NAME, string vORG_NAME)
        {
            if (pSAVE_FileName != string.Empty && pFTP_FILE_NAME != string.Empty && vSAVE_NAME != string.Empty)
            {
                if (DownLoadFile(pFILE_ENTRY_ID, pSAVE_FileName, pFTP_FILE_NAME, vSAVE_NAME, vORG_NAME) == true)
                {
                    return string.Format("{0}", pSAVE_FileName);
                }
                else
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), isMessageAdapter1.ReturnText("EAPP_10206"));
                    return string.Empty;
                }
            }
            else
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), isMessageAdapter1.ReturnText("EAPP_10206"));
                return string.Empty;
            }
        }

        #endregion

        #region ----- file Download Methods -----
        //ftp file download 처리 
        private bool DownLoadFile(object pQUALIFICATION_ID, string pSAVE_FileName, string pFTP_FILE_NAME, string vSAVE_NAME, string vORG_NAME)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool IsDownload = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            //1. 첨부파일 로그 저장 : Transaction을 이용해서 처리 
            //IDC_INSERT_DOC_ATTACHMENT.DataTransaction = isDataTransaction1;
            //IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = isDataTransaction1;

            //isDataTransaction1.BeginTran();
            //IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", pQUALIFICATION_ID);
            //IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "OUT");
            //IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();

            //vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            //vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));

            //if (vSTATUS == "F")
            //{
            //    isDataTransaction1.RollBack();
            //    IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
            //    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}

            //2. 실제 다운로드
            string vTempFileName = string.Format("_{0}", pFTP_FILE_NAME);
            string vClientFileName = string.Format("{0}", pSAVE_FileName);


            //try
            //{
            //FileTransfer Initialze
            mFileTransfer = new ISFileTransferAdv();
            mFileTransfer.Host = mFTP_Info.Host;
            mFileTransfer.Port = mFTP_Info.Port;
            mFileTransfer.UserId = mFTP_Info.UserID;
            mFileTransfer.Password = mFTP_Info.Password;
            if (mFTP_Info.Passive == "Y")
            {
                mFileTransfer.UsePassive = true;
            }
            else
            {
                mFileTransfer.UsePassive = false;
            }

            mFileTransfer.ShowProgress = false;      //진행바 보이기 

            string vSelectFullPath = saveFileDialog1.FileName;
            string vSelectDirectoryPath = Path.GetDirectoryName(saveFileDialog1.FileName);

            string vFileName = Path.GetFileName(saveFileDialog1.FileName);
            string vFileExtension = Path.GetExtension(saveFileDialog1.FileName).ToUpper();

            string fileNameFath = saveFileDialog1.FileName;
            string filePath = fileNameFath.Replace(vFileName, "");

            string vSourceDownLoadFile = string.Format("{0}{1}", filePath, vFileName);
            string vTargetDownLoadFile = string.Format("{0}_{1}", filePath, vFileName);

            string vBeforeSourceFileName = string.Format("{0}", vFileName);
            string vBeforeTargetFileName = string.Format("_{0}", vFileName);

            //다운로드 환경 설정 
            //mFileTransferAdv.SourceDiretory = mFTP_BaseWorkingDiretory;
            mFileTransfer.SourceDirectory = mFTP_Info.FTP_Folder;
            mFileTransfer.SourceFileName = pFTP_FILE_NAME;
            mFileTransfer.TargetDirectory = filePath;
            mFileTransfer.TargetFileName = vBeforeTargetFileName;

            IsDownload = mFileTransfer.Download();

            if (IsDownload == true)
            {
                try
                {
                    //다운 파일 FullPath적용 
                    string vTempFullPath = string.Format("{0}\\{1}", filePath, vBeforeTargetFileName);      //임시
                    string vClientFullPath = string.Format("{0}\\{1}", filePath, vSAVE_NAME);  //원본

                    System.IO.File.Delete(vClientFullPath);                 //기존 파일 삭제 
                    System.IO.File.Move(vTempFullPath, vClientFullPath);    //ftp 이름으로 이름 변경 

                    IsDownload = true;

                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("NFKEAPP_10224"), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch
                {
                    //isDataTransaction1.RollBack();
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTempFileName);
                            }
                            catch
                            {
                                // ignore
                            }
                        }
                    }
                    catch
                    {
                        //ignore                        
                    }
                }
            }
            else
            {
                //isDataTransaction1.RollBack();
                //download 실패 
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTempFileName);
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
                catch
                {
                    //ignore                    
                }
            }

            //isDataTransaction1.Commit();

            //IDA_VENDOR.Fill();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            return IsDownload;
        }

        #endregion;


        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pEdit.PromptTextElement[mIDX].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                    break;
            }
            return mPrompt;
        }

        #endregion;

        #region ---- Form Event -----

        private void H_ZIP_CODE_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address();
            }
        }

        private void H_ADDRESS1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address();
            }
        }

        private void V_VENDOR_DESC_KeyUp(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                AR_CUSTOMER_INQUIRY();
                // Select Error Routine
                if (IDA_VENDOR.GetSelectParamValue("X_ERR_MSG") != null)
                {
                    MessageBoxAdv.Show(IDA_VENDOR.GetSelectParamValue("X_ERR_MSG").ToString());
                }
            }
        }

        private void CUST_SITE_CODE_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (IDA_VENDOR.CurrentRow != null && IDA_VENDOR.CurrentRow.RowState == DataRowState.Added)
            {
                string V_Check_Result = null;

                idcCHK_CUST_SITE_CODE_DUP.ExecuteNonQuery();
                V_Check_Result = idcCHK_CUST_SITE_CODE_DUP.GetCommandParamValue("X_CHECK_RESULT").ToString();


                if (V_Check_Result == 'N'.ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90003", "&&FIELD_NAME:=Customer Site Code"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }
            }
        }

        private void TAX_REG_NO_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (IDA_VENDOR.CurrentRow != null && IDA_VENDOR.CurrentRow.RowState == DataRowState.Added)
            {
                string V_Check_Result = null;

                idcCHK_TAX_REG_NO_DUP.ExecuteNonQuery();
                V_Check_Result = idcCHK_TAX_REG_NO_DUP.GetCommandParamValue("X_CHECK_RESULT").ToString();

                if (V_Check_Result == "N".ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90003", "&&FIELD_NAME:=Tax Registration No"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }
            }
        }

        private void V_RD_ALL_CheckChanged(object sender, EventArgs e)
        {
            if (V_RD_ALL.Checked == true)
            {
                V_VENDOR_TYPE.EditValue = V_RD_ALL.CheckedString.ToString();
            }
        }

        private void V_RD_CUSTOMER_CheckChanged(object sender, EventArgs e)
        {
            if (V_RD_CUSTOMER.Checked == true)
            {
                V_VENDOR_TYPE.EditValue = V_RD_CUSTOMER.CheckedString.ToString();
            }
        }

        private void V_RD_SUPPLIER_CheckChanged(object sender, EventArgs e)
        {
            if (V_RD_SUPPLIER.Checked == true)
            {
                V_VENDOR_TYPE.EditValue = V_RD_SUPPLIER.CheckedString.ToString();
            }
        }

        private void FILE_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (ISG_VENDOR.RowIndex < 0)
            {
                return;
            }

            if (iConv.ISNull(ISG_VENDOR.GetCellValue("VENDOR_ID")) == string.Empty)
            {
                return;
            }
            string Plus_Name = "S";
            if (UpLoadFile(ISG_VENDOR.GetCellValue("VENDOR_ID"), ISG_VENDOR.GetCellValue("VENDOR_CODE"), Plus_Name) == false)
            {
                isAppInterfaceAdv1.OnAppMessage("File Upload fail. Retry it");
                return;
            }
            SEARCH_DB_ATTACHMENT(ISG_VENDOR.GetCellValue("VENDOR_ID"));
        }

        private void FILE_DOWNLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            object vFILE_ENTRY_ID = ISG_ATTACH_FILE.GetCellValue("DOC_ATTACHMENT_ID");
            string vFTP_FILE_NAME = iConv.ISNull(ISG_ATTACH_FILE.GetCellValue("FTP_FILE_NAME"));
            string vEXTENSION_NAME = iConv.ISNull(ISG_ATTACH_FILE.GetCellValue("EXTENSION_NAME"));
            string vFormat = string.Format("{0}{1}", "All file(*.*) | ", vEXTENSION_NAME);

            // 저장될 Dialog 열기
            saveFileDialog1.Title = "Select Save Folder";
            saveFileDialog1.FileName = iConv.ISNull(ISG_ATTACH_FILE.GetCellValue("USER_FILE_NAME"));
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = "C:\\";
            //saveFileDialog2.Filter = "All File(*.*)|*.*|Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|PowerPoint File(*.ppt;*.pptx)|*.ppt;*.pptx|jpg file(*.jpg)|*.jpg";
            saveFileDialog1.Filter = vFormat; 
            if (ISG_ATTACH_FILE.RowIndex < 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("NFKEAPP_10225"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                string vORIGIN_FILE_NAME = iConv.ISNull(ISG_ATTACH_FILE.GetCellValue("USER_FILE_NAME"));
                string vSAVE_NAME = Path.GetFileName(saveFileDialog1.FileName);
                string vSAVE_FILE_NAME = saveFileDialog1.FileName;

                try
                {
                    isDownload(vFILE_ENTRY_ID, vSAVE_FILE_NAME, vFTP_FILE_NAME, vSAVE_NAME, vORIGIN_FILE_NAME);

                    //MessageBoxAdv.Show("다운로드가 완료되었습니다.", "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10212"), isMessageAdapter1.ReturnText("EAPP_10206"));
                    MessageBox.Show("Error : Could not read file from disk.");
                }
            }
        }

        private void FILE_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            object vDOC_ATTACHMENT_ID = ISG_ATTACH_FILE.GetCellValue("DOC_ATTACHMENT_ID");
            if (iConv.ISNull(vDOC_ATTACHMENT_ID) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            IDC_DELETE_DOC_ATTACHMENT.SetCommandParamValue("W_DOC_ATTACHMENT_ID", vDOC_ATTACHMENT_ID);
            IDC_DELETE_DOC_ATTACHMENT.ExecuteNonQuery();

            string vSTATUS = iConv.ISNull(IDC_DELETE_DOC_ATTACHMENT.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_DELETE_DOC_ATTACHMENT.GetCommandParamValue("O_MESSAGE"));

            if (IDC_DELETE_DOC_ATTACHMENT.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            SEARCH_DB_ATTACHMENT(H_VENDOR_ID.EditValue);
        }

        private void BTN_TAX_REG_NUM_CHK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Tax_Reg_Num_Status();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaCUST_PARTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCUST_PARTY.SetLookupParamValue("W_LOOKUP_MODULE", Convert.ToString("PO"));
            ildCUST_PARTY.SetLookupParamValue("W_LOOKUP_TYPE", Convert.ToString("SUPPLIER_CLASS"));
        }

        private void ilaBUSINESS_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "BUSINESS_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_COUNTRY_SelectedRowData(object pSender)
        {
            if(iConv.ISNull(H_COUNTRY_CODE.EditValue) == "KR")
            {
                Show_Tax_Reg_Num(true);
            }
            else
            {
                Show_Tax_Reg_Num(false);
            }
        }

        private void ilaSHIPPING_METHOD_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SHIPPING_METHOD.SetLookupParamValue("W_LOOKUP_MODULE", Convert.ToString("EAPP"));
            ILD_SHIPPING_METHOD.SetLookupParamValue("W_LOOKUP_TYPE", Convert.ToString("SHIPPING_METHOD"));
        }

        private void ilaTAX_BILL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildTAX_BILL_TYPE.SetLookupParamValue("W_LOOKUP_MODULE", Convert.ToString("EAPP"));
            ildTAX_BILL_TYPE.SetLookupParamValue("W_LOOKUP_TYPE", Convert.ToString("TAX_BILL_TYPE"));
        }

        private void ilaACCOUNT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_TYPE");
            //ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", " VALUE1 = 'Y' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaBANK_SITE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildBANK_GROUP.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_BANK_ACCOUNT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OWNER_BANK_ACCOUNT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_TR_ITEM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_TR_ITEM.SetLookupParamValue("W_ENTRY_FLAG", "Y");
            ILD_TR_ITEM.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_VAT_TAX_TYPE_AP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VAT_TAX_TYPE_AP.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_VAT_TAX_TYPE_AR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VAT_TAX_TYPE_AR.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        //private void ilaCURRENCY_BANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        //{
        //    ildCURRENCY_BANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        //}

        //private void ilaBANK_SITE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        //{
        //    ildBANK_SITE.SetLookupParamValue("W_ENABLED_YN", "Y");
        //}

        #endregion

        #region ----- Adapter Event -----

        private void IDA_VENDOR_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["VENDOR_FULL_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(H_VENDOR_FULL_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["VENDOR_SHORT_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(H_VENDORE_SHORT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["CUSTOMER_FLAG"]) == "N" && iConv.ISNull(e.Row["SUPPLIER_FLAG"]) == "N")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10589"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["CUSTOMER_FLAG"]) == "Y" && iConv.ISNull(e.Row["VENDOR_SHORT_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("ISOE_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_VENDOR_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if(pBindingManager.DataRow == null)
            {
                Show_Tax_Reg_Num(false);
                return;
            }
            if (iConv.ISNull(pBindingManager.DataRow["COUNTRY_CODE"]) == "KR")
            {
                Show_Tax_Reg_Num(true);
            }
            else
            {
                Show_Tax_Reg_Num(false);
            }
        }

        private void idaCUST_BANK_ACCT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["BANK_ACCOUNT_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Bank Account Name(은행 계좌명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //코드 입력
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["BANK_ACCOUNT_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Bank Account Number(은행 계좌번호)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["OWNER_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Owner Name(예금주)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["ACCOUNT_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Account Type(계좌 종류)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["CURRENCY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Currency Code(통화)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["EFFECTIVE_DATE_FR"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 시작일자 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["EFFECTIVE_DATE_TO"] != DBNull.Value)
            {
                if (Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
                {// 시작일자 ~ 종료일자
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 기간 검증 오류
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void idaCUST_BANK_ACCT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10047"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        #endregion


        #region ---- 사업자상태 조회 -----

        private string Tax_Reg_Num_Status()
        {
            string vStatus = string.Empty;
            if (iConv.ISNull(H_TAX_REG_NO.EditValue) != string.Empty)
            {
                CRN_Dotnet.CrnWeb vWeb = new CRN_Dotnet.CrnWeb();
                vStatus = vWeb.postCRN(iConv.ISNull(H_TAX_REG_NO.EditValue).Replace("-", ""));
                MessageBoxAdv.Show(vStatus, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return vStatus;
        }
        #endregion

    }

    #region ----- 사업자상태 조회 ------

    namespace CRN_Dotnet
    {
        class CrnWeb
        {
            string vStatus = string.Empty;
            private readonly String postUrl = "https://teht.hometax.go.kr/wqAction.do?actionId=ATTABZAA001R08&screenId=UTEABAAA13&popupYn=false&realScreenId=";
            private String xmlRaw = "<map id=\"ATTABZAA001R08\"><pubcUserNo/><mobYn>N</mobYn><inqrTrgtClCd>1</inqrTrgtClCd><txprDscmNo>{CRN}</txprDscmNo><dongCode>15</dongCode><psbSearch>Y</psbSearch><map id=\"userReqInfoVO\"/></map>";

            public String postCRN(String crn)
            {
                byte[] contents = System.Text.Encoding.ASCII.GetBytes(xmlRaw.Replace("{CRN}", crn));
                HttpWebRequest request = createHttpWebRequest();
                setContentStream(request, contents);

                HttpWebResponse response;
                response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    Stream responseStream = response.GetResponseStream();
                    String resString = new StreamReader(responseStream).ReadToEnd();
                    responseStream.Close();
                    response.Close();
                    vStatus = getCRNresultFromXml(resString);
                }
                response.Close();
                return vStatus;
            }
            private HttpWebRequest createHttpWebRequest()
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(postUrl);
                request.ContentType = "text/xml; encoding='utf-8'";
                request.Method = "POST";
                return request;
            }

            private void setContentStream(HttpWebRequest request, byte[] contents)
            {
                request.ContentLength = contents.Length;
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(contents, 0, contents.Length);
                requestStream.Close();
            }

            private String getCRNresultFromXml(String xmlData)
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlData);
                String crnResult = xmlDocument.SelectNodes("//trtCntn").Item(0).InnerText;
                return crnResult;
            }
        }
    }

    #endregion

    #region ----- FTP 정보 위한 사용자 Class -----

    public class isFTP_Info
    {
        #region ----- Variables -----

        private string mHost = string.Empty;
        private string mPort = string.Empty;
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;
        private string mFTP_Folder = string.Empty;
        private string mClient_Folder = string.Empty;
        private string mPassive = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public isFTP_Info()
        {

        }

        public isFTP_Info(string pHost, string pPort, string pUserID, string pPassword, string pFTP_Folder, string pClient_Folder, string pPassive)
        {
            mHost = pHost;
            mPort = pPort;
            mUserID = pUserID;
            mPassword = pPassword;
            mFTP_Folder = pFTP_Folder;
            mClient_Folder = pClient_Folder;
            mPassive = pPassive;
        }

        #endregion;

        #region ----- Property -----

        public string Host
        {
            get
            {
                return mHost;
            }
            set
            {
                mHost = value;
            }
        }

        public string Port
        {
            get
            {
                return mPort;
            }
            set
            {
                mPort = value;
            }
        }

        public string UserID
        {
            get
            {
                return mUserID;
            }
            set
            {
                mUserID = value;
            }
        }

        public string Password
        {
            get
            {
                return mPassword;
            }
            set
            {
                mPassword = value;
            }
        }

        public string FTP_Folder
        {
            get
            {
                return mFTP_Folder;
            }
            set
            {
                mFTP_Folder = value;
            }
        }

        public string Client_Folder
        {
            get
            {
                return mClient_Folder;
            }
            set
            {
                mClient_Folder = value;
            }
        }

        public string Passive
        {
            get
            {
                return mPassive;
            }
            set
            {
                mPassive = value;
            }
        }

        #endregion;
    }

    #endregion

}