using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections.Generic;
using System.ComponentModel;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
using System.IO;

namespace EAPF0411
{
    public partial class EAPF0411 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- UpLoad / DownLoad Variables -----

        private InfoSummit.Win.ControlAdv.ISFileTransferAdv mFileTransferAdv;
        private ItemImageInfomationFTP mImageFTP;

        private string mFTP_Source_Directory = string.Empty;            // ftp 소스 디렉토리.
        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 디렉토리.
        private string mClient_Target_Directory = string.Empty;         // 실제 디렉토리 
        private string mClient_ImageDirectory = string.Empty;           // 클라이언트 이미지 디렉토리.
        private string mFileExtension = string.Empty;                   // 확장자명.

        private bool mIsGetInformationFTP = false;                      // FTP 정보 상태.
        private bool mIsFormLoad = false;                               // NEWMOVE 이벤트 제어.

        #endregion; 

        #region ----- Constructor -----

        public EAPF0411()
        {
            InitializeComponent();
        }

        public EAPF0411(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mIsFormLoad = false;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            if (iConv.ISNull(V_NOTICE_DATE_FR.EditValue) == String.Empty)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_NOTICE_DATE_FR.Focus();
                return;
            }
            if (iConv.ISNull(V_NOTICE_DATE_TO.EditValue) == String.Empty)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_NOTICE_DATE_TO.Focus();
                return;
            }
            if (Convert.ToDateTime(V_NOTICE_DATE_FR.EditValue) > Convert.ToDateTime(V_NOTICE_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_NOTICE_DATE_FR.Focus();
                return;
            }
            
            if (TB_MAIN.SelectedTab.TabIndex == 1)
            {
                

                //IGR_NOTICE_ALL.Focus();
                string vNOTICE_ID = iConv.ISNull(IGR_FILE_LIST.GetCellValue("NOTICE_ID"));
                int vIDX_Col = IGR_NOTICE_ALL.GetColumnToIndex("NOTICE_ID");

                IDA_NOTICE_ALL.Fill();

                if (IGR_NOTICE_ALL.RowCount > 0)
                {
                    for (int vRow = 0; vRow < IGR_NOTICE_ALL.RowCount; vRow++)
                    {
                        if (vNOTICE_ID == iConv.ISNull(IGR_NOTICE_ALL.GetCellValue(vRow, vIDX_Col)))
                        {
                            IGR_NOTICE_ALL.CurrentCellActivate(vRow, vIDX_Col);
                            IGR_NOTICE_ALL.CurrentCellMoveTo(vRow, vIDX_Col);
                        }
                    }
                }

            }
            else if (TB_MAIN.SelectedTab.TabIndex == 2)
            {
                
            }
        }


        private void INSERT_NOTICE_ALL()
        {
            try
            {
                NOTICE_DATE.EditValue = DateTime.Today;

                IDC_GET_DEFAULT_NOTICE_LEVEL.ExecuteNonQuery();
                NOTICE_LEVEL.EditValue = IDC_GET_DEFAULT_NOTICE_LEVEL.GetCommandParamValue("O_NOTICE_LEVEL");
                NOTICE_LEVEL_DESC.EditValue = IDC_GET_DEFAULT_NOTICE_LEVEL.GetCommandParamValue("O_NOTICE_LEVEL_DESC");

                ENABLED_FLAG.CheckedState = ISUtil.Enum.CheckedState.Checked;
            }
            catch
            {
            }
            NOTICE_DATE.Focus();
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_NOTICE_ALL.IsFocused)
                    {
                        IDA_NOTICE_ALL.AddOver();
                        INSERT_NOTICE_ALL();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_NOTICE_ALL.IsFocused)
                    {
                        IDA_NOTICE_ALL.AddUnder();
                        INSERT_NOTICE_ALL();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_NOTICE_ALL.IsFocused)
                    {
                        IDA_NOTICE_ALL.Update();
                    }
                    if (IDA_FILE_LIST.IsFocused)
                    {

                        IDA_FILE_LIST.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_NOTICE_ALL.IsFocused)
                    {
                        IDA_NOTICE_ALL.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_NOTICE_ALL.IsFocused)
                    {
                        IDA_NOTICE_ALL.Delete();
                    }
                    if (IDA_FILE_LIST.IsFocused)
                    {
                        
                       IDA_FILE_LIST.Delete();
                        
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void EAPF0411_Load(object sender, EventArgs e)
        {
            V_NOTICE_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            V_NOTICE_DATE_TO.EditValue = DateTime.Today;
        }

        private void EAPF0411_Shown(object sender, EventArgs e)
        {
            IDA_NOTICE_ALL.FillSchema();

            //DefaultCorporation();
            mIsGetInformationFTP = GetInfomationFTP();
            if (mIsGetInformationFTP == true)
            {
                MakeDirectory();
                FTPInitializtion();
            }
            mIsFormLoad = false;
        }

        private void ibtUPLOAD_CSR_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string vFileName = iConv.ISNull(IGR_NOTICE_ALL.GetCellValue("NOTICE_ID"));

            if (iConv.ISNull(IGR_NOTICE_ALL.GetCellValue("NOTICE_ID")) == "0")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10081"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (mIsGetInformationFTP == true)
            {
                UpLoadFile();
            } 
        }

        #endregion

        #region ----- Make Directory ----

        private void MakeDirectory()
        {
            System.IO.DirectoryInfo vClient_ImageDirectory = new System.IO.DirectoryInfo(mClient_ImageDirectory);
            if (vClient_ImageDirectory.Exists == false) //있으면 True, 없으면 False
            {
                vClient_ImageDirectory.Create();
            }
        }

        #endregion;

        #region ----- Get Information FTP Methods -----

        private bool GetInfomationFTP()
        {
            bool isGet = false;
            try
            {
                idcFTP_INFO.SetCommandParamValue("W_FTP_INFO_CODE", "30");
                idcFTP_INFO.ExecuteNonQuery();
                mImageFTP = new ItemImageInfomationFTP();

                mImageFTP.Host = iConv.ISNull(idcFTP_INFO.GetCommandParamValue("O_FTP_IP"));
                mImageFTP.Port = iConv.ISNull(idcFTP_INFO.GetCommandParamValue("O_FTP_PORT"));
                mImageFTP.UserID = iConv.ISNull(idcFTP_INFO.GetCommandParamValue("O_FTP_USER_ID"));
                mImageFTP.Password = iConv.ISNull(idcFTP_INFO.GetCommandParamValue("O_FTP_PASSWORD"));

                mFTP_Source_Directory = iConv.ISNull(idcFTP_INFO.GetCommandParamValue("O_FTP_SOURCEPATH"));
                mClient_Target_Directory = iConv.ISNull(idcFTP_INFO.GetCommandParamValue("O_CLIENT_TARGETPATH"));
                mFileExtension = ".JPG";
                mClient_ImageDirectory = string.Format("{0}\\{1}", mClient_Base_Path, mClient_Target_Directory);

                Application.DoEvents();

                if (mImageFTP.Host != string.Empty)
                {
                    isGet = true;
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return isGet;
        }

        #endregion;

        #region ----- FTP Initialize -----

        private void FTPInitializtion()
        {
            mFileTransferAdv = new ISFileTransferAdv();
            mFileTransferAdv.Host = mImageFTP.Host;
            mFileTransferAdv.Port = mImageFTP.Port;
            mFileTransferAdv.UserId = mImageFTP.UserID;
            mFileTransferAdv.Password = mImageFTP.Password;
        }

        #endregion;

        #region ----- File Upload Methods -----

        private bool UpLoadFile()
        {
            bool isUp = false;

            if (iConv.ISNull(IGR_NOTICE_ALL.GetCellValue("NOTICE_ID")) != "")
            {
                //bool isUp = false;
                string vFileExtension = Path.GetExtension(openFileDialog1.FileName);

                //openFileDialog1.FileName = string.Format("*{0}", vFileExtension);
                //openFileDialog1.Filter = string.Format("Image Files (*{0})|*{1}", vFileExtension, vFileExtension);

                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|jpg file(*.jpg)|*.jpg|All File(*.*)|*.*";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "";
                openFileDialog1.Multiselect = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string vChoiceFileFullPath = string.Empty;
                        int vArryCount = openFileDialog1.FileNames.Length;
                        for (int r = 0; r < vArryCount; r++)
                        {

                            vChoiceFileFullPath = openFileDialog1.FileNames.GetValue(r).ToString();


                            string vChoiceFilePath = vChoiceFileFullPath.Substring(0, vChoiceFileFullPath.LastIndexOf(@"\"));
                            string vChoiceFileName = vChoiceFileFullPath.Substring(vChoiceFileFullPath.LastIndexOf(@"\") + 1);
                            vFileExtension = Path.GetExtension(openFileDialog1.FileName);
                            mFileTransferAdv.ShowProgress = true;
                            //--------------------------------------------------------------------------------
                            IDC_FILE_FTPNAME.ExecuteNonQuery();

                            string vTargetFileName = FTP_FILE_NAME.EditValue as string;
                            vTargetFileName = string.Format("{0}{1}", vTargetFileName.ToUpper(), vFileExtension);

                            mFileTransferAdv.SourceDirectory = vChoiceFilePath;
                            mFileTransferAdv.SourceFileName = vChoiceFileName;
                            mFileTransferAdv.TargetDirectory = mFTP_Source_Directory;
                            mFileTransferAdv.TargetFileName = vTargetFileName;

                            bool isUpLoad = mFileTransferAdv.Upload();

                            if (isUpLoad == true)
                            {
                                isUp = true;
                                FTP_FILE_NAME.EditValue = vTargetFileName;
                                FILE_NAME.EditValue = vChoiceFileName;

                                IDC_FILE_UPLOAD_BTN_CLICK.ExecuteNonQuery();
                                SEARCH_DB();


                            }
                            else
                            {
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10076"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            }

                        }




                    }
                    catch
                    {
                    }
                }
                System.IO.Directory.SetCurrentDirectory(mClient_Base_Path);
                return isUp;
            }
            else
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10082"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return isUp;
            }

        }
        #endregion;

        #region ----- file Download Methods -----

        private bool DownLoadFile(string pFileName)
        {

            bool isDown = false;

            string vSourceDownLoadFile = string.Format("{0}\\{1}", mClient_ImageDirectory, pFileName);
            string vTargetDownLoadFile = string.Format("{0}\\_{1}", mClient_ImageDirectory, pFileName);

            string vBeforeSourceFileName = string.Format("{0}", pFileName);
            string vBeforeTargetFileName = string.Format("_{0}", pFileName);

            mFileTransferAdv.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransferAdv.SourceDirectory = mFTP_Source_Directory;
            mFileTransferAdv.SourceFileName = vBeforeSourceFileName;
            mFileTransferAdv.TargetDirectory = mClient_ImageDirectory;
            mFileTransferAdv.TargetFileName = vBeforeTargetFileName;

            isDown = mFileTransferAdv.Download();

            if (isDown == true)
            {
                try
                {
                    System.IO.File.Delete(vSourceDownLoadFile);
                    System.IO.File.Move(vTargetDownLoadFile, vSourceDownLoadFile);

                    isDown = true;
                }
                catch
                {
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTargetDownLoadFile);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTargetDownLoadFile);
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
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTargetDownLoadFile);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTargetDownLoadFile);
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

            return isDown;
        }

        #endregion;  

        #region ----- is View Measure file Method -----

        private string isViewMeasureFile(object pFileName)
        {
            if (iConv.ISNull(pFileName) != string.Empty)
            {
                string vFileName = iConv.ISNull(pFileName);

                string vDownLoadFile = string.Empty;

                bool isDown = DownLoadFile(vFileName);

                if (isDown == true)
                {
                    return string.Format("{0}\\{1}", mClient_ImageDirectory, vFileName);
                }
            }
            return string.Empty;
        }

        #endregion;

        private void IDA_NOTICE_ALL_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            IDA_FILE_LIST.SetSelectParamValue("W_NOTICE_ID", pBindingManager.DataRow["NOTICE_ID"]);
            IDA_FILE_LIST.Fill();
        }

        private void IGR_FILE_LIST_CellDoubleClick(object pSender)
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            try
            {
                string vFileName = iConv.ISNull(IGR_FILE_LIST.GetCellValue("FILE_CODE"));

                vFileName = isViewMeasureFile(vFileName);

                Application.UseWaitCursor = false;
                this.Cursor = System.Windows.Forms.Cursors.Default;
                Application.DoEvents();

                if (vFileName == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //vFileName = string.Format("{0}\\{1}", mClient_ImageDirectory, vFileName);
                System.Diagnostics.Process.Start(vFileName);
            }
            catch (Exception ex)
            {
                Application.UseWaitCursor = false;
                this.Cursor = System.Windows.Forms.Cursors.Default;
                Application.DoEvents();

                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }


        
         
    }
    #region ----- User Make Class -----

    public class ItemImageInfomationFTP
    {
        #region ----- Variables -----

        private string mHost = string.Empty;
        private string mPort = string.Empty;
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public ItemImageInfomationFTP()
        {
        }

        public ItemImageInfomationFTP(string pHost, string pPort, string pUserID, string pPassword)
        {
            mHost = pHost;
            mPort = pPort;
            mUserID = pUserID;
            mPassword = pPassword;
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

        #endregion;
    }

    #endregion;
}