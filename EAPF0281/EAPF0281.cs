using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
using System.Net;
using System.Xml;
using Syncfusion.XlsIO.Parser.Biff_Records;
using System.IO;
using System.Text; 


namespace EAPF0281
{
    public partial class EAPF0281 : Office2007Form
    {
        #region ----- Variables -----
          
        string mApiUrl = string.Empty;

        ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        int mTIME_SEC = 0;

        #endregion;

        #region ----- Constructor -----

        public EAPF0281()
        {
            InitializeComponent();
        }

        public EAPF0281(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm; //항상 최상위 폼 유지 위해.
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
         
        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            IDA_BANK_IF_EXCH_RATE.Fill();
            IGR_BANK_IF_EXCH_RATE.Focus();
        }
         
        private void Init_Timer(bool pEnabled_Flag)
        {
            if(iConvert.ISNull(W_SYNC_CYCLE_TIME.EditValue) == "10")
            {
                V_TIMER.Interval = 1000 * 60 * 10;
            }
            else if (iConvert.ISNull(W_SYNC_CYCLE_TIME.EditValue) == "5")
            {
                V_TIMER.Interval = 1000 * 60 * 5;
            }
            else
            {
                V_TIMER.Interval = 1000 * 60 * 3;
            }
            if (pEnabled_Flag == true)
            {
                V_TIMER.Enabled = true;
                V_TIMER_SEC.Enabled = true; 
            }
            else
            {
                V_TIMER.Enabled = false;
                V_TIMER_SEC.Enabled = false;
            }
            mTIME_SEC = 1;
        }

        private void Get_Exch_Rate_IF_Bank()
        {
            IDC_GET_EXCH_RATE_IF_BANK_DFT.ExecuteNonQuery();
            if (iConvert.ISNull(IDC_GET_EXCH_RATE_IF_BANK_DFT.GetCommandParamValue("O_BANK_TYPE")) != string.Empty)
            {
                V_BANK_TYPE.EditValue = IDC_GET_EXCH_RATE_IF_BANK_DFT.GetCommandParamValue("O_BANK_TYPE");
                V_BANK_TYPE_NAME.EditValue = IDC_GET_EXCH_RATE_IF_BANK_DFT.GetCommandParamValue("O_BANK_TYPE_NAME");
            }
        }

        private bool Sync_Exch_Rate()
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            //message
            V_MESSAGE.PromptText = "Timer 중지 중....";
            Application.DoEvents();
              
            if (iConvert.ISNull(V_BANK_TYPE.EditValue) == "KEB")
            {
                if(Sync_Exch_Rate_KEB() == false)
                {
                    return false;
                }
            }
            else
            {
                if (Sync_Exch_Rate_EIBK() == false)
                {
                    return false;
                }
            } 

            V_MESSAGE.PromptText = "Exchange Rate 정보 수신을 완료했습니다.";
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            return true;
        }

        private bool Sync_Exch_Rate_EIBK()
        {             
            V_MESSAGE.PromptText = "Host 정보 수신중 중....";
            Application.DoEvents();

            IDC_GET_EXCH_RATE_IF_BANK.ExecuteNonQuery();
            V_BANK_TYPE_NAME.EditValue = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_BANK_TYPE_NAME");
            object vHOST = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_HOST");
            object vPORT = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PORT");
            object vUSER_ID = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_USER_ID");
            object vPWD = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PWD");
            object vSSL_FLAG = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_SSL_FLAG");
            object vAUTH_KEY = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_AUTH_KEY");
            object vPARAMETER1 = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PARAMETER1");
            object vPARAMETER2 = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PARAMETER2");
            object vPARAMETER3 = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PARAMETER3");
            object vPARAMETER4 = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PARAMETER4");
            V_EXCH_RATE_IF_BANK.PromptText = String.Format("{0} :: {1}", V_BANK_TYPE_NAME.EditValue, vHOST);
            if (iConvert.ISNull(vHOST) == string.Empty)
            {
                V_EXCH_RATE_IF_BANK.PromptText = String.Empty;
                return false;
            }

            V_MESSAGE.PromptText = "Exchange Rate 정보 수신중 중....";
            Application.DoEvents();

            mApiUrl = iConvert.ISNull(vHOST);

            StringBuilder dsDataString = new StringBuilder();
            if (iConvert.ISNull(vPARAMETER1) != String.Empty)
            {
                dsDataString.Append(vPARAMETER1);
            }
            if (iConvert.ISNull(vPARAMETER2) != String.Empty)
            {
                dsDataString.Append(vPARAMETER2);
            }
            if (iConvert.ISNull(vPARAMETER3) != String.Empty)
            {
                dsDataString.Append(vPARAMETER3);
            }
            if (iConvert.ISNull(vPARAMETER4) != String.Empty)
            {
                dsDataString.Append(vPARAMETER4);
            }

            // 요청 String->요청 Byte 변환
            byte[] byteDataParams = System.Text.UTF8Encoding.UTF8.GetBytes(dsDataString.ToString());
             
            /* POST */
            // HttpWebRequest 객체 생성, 설정
            HttpWebRequest reqRequest = (HttpWebRequest)WebRequest.Create(mApiUrl);
            reqRequest.Method = "POST";  // 기본값 "GET"
            reqRequest.ContentType = "application/x-www-form-urlencoded";
            reqRequest.ContentLength = byteDataParams.Length;

            /* GET */
            // GET 방식은 Uri 뒤에 보낼 데이터를 입력하시면 됩니다.

            //HttpWebRequest reqRequest = (HttpWebRequest)WebRequest.Create(mApiUrl);
            //reqRequest.Method = "GET"; 

            //요청 Byte -> 요청 Stream 변환
            Stream stData = reqRequest.GetRequestStream();
            stData.Write(byteDataParams, 0, byteDataParams.Length);
            stData.Close();

            try
            {
                // 요청, 응답 받기
                HttpWebResponse rsResponse = (HttpWebResponse)reqRequest.GetResponse();

                //응답 Stream 읽기.
                Stream stReadData = rsResponse.GetResponseStream();
                StreamReader srReadData = new StreamReader(stReadData, System.Text.Encoding.UTF8);

                //응답 Stream -> 응답 String 변환
                string strResult = srReadData.ReadToEnd();

                Console.WriteLine(strResult);

                if (strResult != null)
                {
                    string strOri_Exch_Value = iConvert.ISNull(strResult);
                    string strExch_Value = iConvert.ISNull(strResult).Replace("[", "").Replace("]", "");
                    string[] arrExch_Value = strExch_Value.Split('{');

                    for (int r = 0; r < arrExch_Value.Length; r++)
                    {
                        if (arrExch_Value[r] != string.Empty)
                        {
                            IDC_IMPORT_EXCH_VALUE.SetCommandParamValue("P_EXCH_VALUE", arrExch_Value[r]);
                            IDC_IMPORT_EXCH_VALUE.ExecuteNonQuery();
                            string vSTATUS = iConvert.ISNull(IDC_IMPORT_EXCH_VALUE.GetCommandParamValue("O_STATUS"));
                            string vMESSAGE = iConvert.ISNull(IDC_IMPORT_EXCH_VALUE.GetCommandParamValue("O_MESSAGE"));
                            if (vSTATUS == "F")
                            {
                                PB_INSERT.BarFillPercent = 0;
                                V_MESSAGE.PromptText = vMESSAGE;
                                Application.UseWaitCursor = false;
                                System.Windows.Forms.Cursor.Current = Cursors.Default;
                                Application.DoEvents();
                                return true;
                            }
                        }

                        PB_INSERT.BarFillPercent = (Convert.ToSingle(r) / Convert.ToSingle(arrExch_Value.Length)) * 100F;
                        V_MESSAGE.PromptText = string.Format("Importing :: {0} / {1} = {2}", r, arrExch_Value.Length, arrExch_Value[r]);
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                        Application.DoEvents();
                    }
                    PB_INSERT.BarFillPercent = 0;
                }
            }
            catch (Exception Ex)
            {
                V_MESSAGE.PromptText = string.Format("Exchange Rate 응답 오류입니다. {0}", Ex.Message);
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return true;
            } 
            return true;
        }

        private bool Sync_Exch_Rate_KEB()
        {             
            V_MESSAGE.PromptText = "Host 정보 수신중 중....";
            Application.DoEvents();

            IDC_GET_EXCH_RATE_IF_BANK.ExecuteNonQuery();
            V_BANK_TYPE_NAME.EditValue = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_BANK_TYPE_NAME");
            object vHOST = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_HOST");
            object vPORT = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PORT");
            object vUSER_ID = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_USER_ID");
            object vPWD = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PWD");
            object vSSL_FLAG = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_SSL_FLAG");
            object vAUTH_KEY = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_AUTH_KEY");
            object vPARAMETER1 = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PARAMETER1");
            object vPARAMETER2 = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PARAMETER2");
            object vPARAMETER3 = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PARAMETER3");
            object vPARAMETER4 = IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_PARAMETER4");
            V_EXCH_RATE_IF_BANK.PromptText = String.Format("{0} :: {1}", V_BANK_TYPE_NAME.EditValue, vHOST);
            if (iConvert.ISNull(vHOST) == string.Empty)
            {
                V_EXCH_RATE_IF_BANK.PromptText = String.Empty;
                return false;
            }

            V_MESSAGE.PromptText = "Exchange Rate 정보 수신중 중....";
            Application.DoEvents();

            mApiUrl = iConvert.ISNull(vHOST);

            StringBuilder dsDataString = new StringBuilder();
            if (iConvert.ISNull(vPARAMETER1) != String.Empty)
            {
                dsDataString.Append(vPARAMETER1);
            }
            if (iConvert.ISNull(vPARAMETER2) != String.Empty)
            {
                dsDataString.Append(vPARAMETER2);
            }
            if (iConvert.ISNull(vPARAMETER3) != String.Empty)
            {
                dsDataString.Append(vPARAMETER3);
            }
            if (iConvert.ISNull(vPARAMETER4) != String.Empty)
            {
                dsDataString.Append(vPARAMETER4);
            }

            // 요청 String->요청 Byte 변환
            byte[] byteDataParams = System.Text.UTF8Encoding.UTF8.GetBytes(dsDataString.ToString());


            /* POST */
            // HttpWebRequest 객체 생성, 설정
            HttpWebRequest reqRequest = (HttpWebRequest)WebRequest.Create(mApiUrl);
            reqRequest.Method = "POST";  // 기본값 "GET"
            reqRequest.ContentType = "application/x-www-form-urlencoded";
            reqRequest.ContentLength = byteDataParams.Length;

            /* GET */
            // GET 방식은 Uri 뒤에 보낼 데이터를 입력하시면 됩니다.

            //HttpWebRequest reqRequest = (HttpWebRequest)WebRequest.Create(mApiUrl);
            //reqRequest.Method = "GET"; 

            //요청 Byte -> 요청 Stream 변환
            Stream stData = reqRequest.GetRequestStream();
            stData.Write(byteDataParams, 0, byteDataParams.Length);
            stData.Close();

            try
            {
                // 요청, 응답 받기
                HttpWebResponse rsResponse = (HttpWebResponse)reqRequest.GetResponse();

                //응답 Stream 읽기.
                Stream stReadData = rsResponse.GetResponseStream();
                StreamReader srReadData = new StreamReader(stReadData, System.Text.Encoding.Default, true);

                //응답 Stream -> 응답 String 변환
                string strResult = srReadData.ReadToEnd();

                Console.WriteLine(strResult);

                if (strResult != null)
                {
                    string strOri_Exch_Value = iConvert.ISNull(strResult);
                    string strExch_Value = iConvert.ISNull(strResult).Replace("[", "").Replace("]", "");

                    strExch_Value = iConvert.ISNull(strResult).Replace("var exView =", "").Replace("\r\n", "").Replace("\t", "").Replace("\\", "");
                    string[] arrExch_Value = strExch_Value.Split('{');

                    for (int r = 0; r < arrExch_Value.Length; r++)
                    {
                        Console.WriteLine(arrExch_Value[r]);
                        if (arrExch_Value[r] != string.Empty)
                        {
                            IDC_IMPORT_EXCH_VALUE.SetCommandParamValue("P_EXCH_VALUE", arrExch_Value[r]);
                            IDC_IMPORT_EXCH_VALUE.ExecuteNonQuery();
                            string vSTATUS = iConvert.ISNull(IDC_IMPORT_EXCH_VALUE.GetCommandParamValue("O_STATUS"));
                            string vMESSAGE = iConvert.ISNull(IDC_IMPORT_EXCH_VALUE.GetCommandParamValue("O_MESSAGE"));
                            if (vSTATUS == "F")
                            {
                                PB_INSERT.BarFillPercent = 0;
                                V_MESSAGE.PromptText = vMESSAGE;
                                Application.UseWaitCursor = false;
                                System.Windows.Forms.Cursor.Current = Cursors.Default;
                                Application.DoEvents();
                                return true;
                            }
                        }

                        PB_INSERT.BarFillPercent = (Convert.ToSingle(r) / Convert.ToSingle(arrExch_Value.Length)) * 100F;
                        V_MESSAGE.PromptText = string.Format("Importing :: {0} / {1} = {2}", r, arrExch_Value.Length, arrExch_Value[r]);
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                        Application.DoEvents();
                    }
                    PB_INSERT.BarFillPercent = 0;
                }
            }
            catch (Exception Ex)
            {
                V_MESSAGE.PromptText = string.Format("Exchange Rate 응답 오류입니다. {0}", Ex.Message);
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return true;
            }
            return true;
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

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_BANK_IF_EXCH_RATE.IsFocused)
                    {
                        IDA_BANK_IF_EXCH_RATE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_BANK_IF_EXCH_RATE.IsFocused)
                    {
                        IDA_BANK_IF_EXCH_RATE.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void EAPF0281_Load(object sender, EventArgs e)
        {

        }

        private void EAPF0281_Shown(object sender, EventArgs e)
        {
            Get_Exch_Rate_IF_Bank();
            RB_10.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_SYNC_CYCLE_TIME.EditValue = RB_10.RadioCheckedString;
            Init_Timer(true); 

            PB_INSERT.BarFillPercent = 0;
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void RB_TIME_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB = sender as ISRadioButtonAdv;

            W_SYNC_CYCLE_TIME.EditValue = RB.RadioCheckedString;
            Init_Timer(true);
        }

        private void V_TIMER_Tick(object sender, EventArgs e)
        {
            //timer 중지
            Init_Timer(false);

            if (Sync_Exch_Rate() == false)
            {
                return;
            }
            Init_Timer(true);
        }

        private void V_TIMER_SEC_Tick(object sender, EventArgs e)
        {
            string vMESSAGE = "*";
            mTIME_SEC++;
            if(mTIME_SEC < 61)
            {
                vMESSAGE = "********************************************************************".Substring(1, mTIME_SEC);
            }
            else
            {
                mTIME_SEC = 1;
                vMESSAGE = "*";
            }
            V_WAIT.PromptText = string.Format("Waiting ( {0}s ) : {1}", mTIME_SEC, vMESSAGE);
        }

        private void BTN_TRANSFER_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();


            IDC_SYNC_EXCHANGE_RATE.ExecuteNonQuery();
            string strSTATUS = iConvert.ISNull(IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_STATUS"));
            string strMESSAGE = iConvert.ISNull(IDC_GET_EXCH_RATE_IF_BANK.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            if (strSTATUS == "F")
            { 
                if(strMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(strMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            } 
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            if (strMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(strMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            SEARCH_DB();
        }

        private void BTN_RECEIVING_EXCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //timer 중지
            Init_Timer(false);

            if(Sync_Exch_Rate() == false)
            {
                return;
            }
            Init_Timer(true);
        }

        #endregion

        #region  ----- Lookup ----- 

        private void ILA_CURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion;

    }
}
 