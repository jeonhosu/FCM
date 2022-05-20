using System;

namespace FCMF0391
{
    public class Zebra
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv mAppInterfaceAdv = null;

        private string mMessageError = string.Empty;

        private System.Text.StringBuilder mLabelString = new System.Text.StringBuilder();

        private System.Data.DataTable mDataTable = null;

        #endregion;

        #region ----- Constructor -----

        public Zebra(InfoSummit.Win.ControlAdv.ISAppInterfaceAdv pAppInterfaceAdv, System.Data.DataTable pDataTable)
        {
            mAppInterfaceAdv = pAppInterfaceAdv;

            mDataTable = pDataTable;
        }

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        #endregion;

        #region ----- Dispose Method -----

        public void Dispose()
        {
        }

        #endregion;

        #region ----- Convert String Methods ----

        private string ConvertString(object pObject)
        {
            string vString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is string;
                    if (IsConvert == true)
                    {
                        vString = pObject as string;
                    }
                }
            }
            catch(System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterfaceAdv.OnAppMessage(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vString;
        }

        #endregion;

        #region ----- Convert Decimal to String Methods ----

        private string ConvertNumberToString(object pObject)
        {
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0;

            try
            {
                if (pObject != null)
                {
                    bool vIsConvert = pObject is decimal;
                    if (vIsConvert == true)
                    {
                        vConvertDecimal = (decimal)pObject;
                        vConvertString = vConvertDecimal.ToString();
                    }
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterfaceAdv.OnAppMessage(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vConvertString;
        }

        #endregion;

        #region ----- Make String ZPL Method -----

        public bool MakeStringZPL(string pManageNo, string pAssetName, string pItemSpec, string pManageF, string pManageS, string pAcquireDate, string pUseDept)
        {
            bool isMake = false;
            object vObject = null;
            int vCountRow = mDataTable.Rows.Count;

            if (vCountRow > 0)
            {
                try
                {
                    //mLabelString = new System.Text.StringBuilder();

                    //------------------------------------------------------------------------------------------------------------
                    // 기준위치
                    //------------------------------------------------------------------------------------------------------------
                    mLabelString.Append("^XA"); //START

                    //mLabelString.Append("^LH30,20"); //POSITION 기준위치[열,행]
                    vObject = mDataTable.Rows[0]["LH_X"];
                    string vLH_X = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["LH_Y"];
                    string vLH_Y = ConvertNumberToString(vObject);
                    mLabelString.Append("^LH").Append(vLH_X).Append(",").Append(vLH_Y); //POSITION 기준위치[열,행]

                    mLabelString.Append("^SEE:UHANGUL.DAT^FS");
                    mLabelString.Append("^CW1.E:KFONT3.FNT^CI26^FS");

                    /*============================================================================================================
                    // Label에 있는 항목이므로 항목 수만큼 반복
                    ============================================================================================================*/
                    //------------------------------------------------------------------------------------------------------------
                    // 관리번호
                    //------------------------------------------------------------------------------------------------------------
                    vObject = mDataTable.Rows[0]["A_FO_X"];
                    string vA_FO_X = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["A_FO_Y"];
                    string vA_FO_Y = ConvertNumberToString(vObject);
                    mLabelString.Append("^FO").Append(vA_FO_X).Append(",").Append(vA_FO_Y); //기준위치에서 지정한 위치에[열,행]

                    //mLabelString.Append("^A1N,36,20"); //Scalable
                    vObject = mDataTable.Rows[0]["A_H"];
                    string vA_H = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["A_W"];
                    string vA_W = ConvertNumberToString(vObject);
                    mLabelString.Append("^A1N,").Append(vA_H).Append(",").Append(vA_W); //Scalable
                    mLabelString.Append("^FD"); //START DATA
                    mLabelString.Append(pManageNo); //출력될 글자
                    mLabelString.Append("^FS"); //END DATA
                    //------------------------------------------------------------------------------------------------------------
                    // 자산명
                    //------------------------------------------------------------------------------------------------------------
                    vObject = mDataTable.Rows[0]["B_FO_X"];
                    string vB_FO_X = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["B_FO_Y"];
                    string vB_FO_Y = ConvertNumberToString(vObject);
                    mLabelString.Append("^FO").Append(vB_FO_X).Append(",").Append(vB_FO_Y); //기준위치에서 지정한 위치에[열,행]

                    //mLabelString.Append("^A1N,36,20"); //Scalable
                    vObject = mDataTable.Rows[0]["B_H"];
                    string vB_H = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["B_W"];
                    string vB_W = ConvertNumberToString(vObject);
                    mLabelString.Append("^A1N,").Append(vB_H).Append(",").Append(vB_W); //Scalable
                    mLabelString.Append("^FD");      //START DATA
                    mLabelString.Append(pAssetName); //출력될 글자
                    mLabelString.Append("^FS");      //END DATA
                    //------------------------------------------------------------------------------------------------------------
                    // 규격
                    //------------------------------------------------------------------------------------------------------------

                    //------------------------------------------------------------------------------------------------------------
                    // 관리자(정)
                    //------------------------------------------------------------------------------------------------------------
                    vObject = mDataTable.Rows[0]["C_FO_X"];
                    string vC_FO_X = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["C_FO_Y"];
                    string vC_FO_Y = ConvertNumberToString(vObject);
                    mLabelString.Append("^FO").Append(vC_FO_X).Append(",").Append(vC_FO_Y); //기준위치에서 지정한 위치에[열,행]

                    //mLabelString.Append("^A1N,36,20"); //Scalable
                    vObject = mDataTable.Rows[0]["C_H"];
                    string vC_H = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["C_W"];
                    string vC_W = ConvertNumberToString(vObject);
                    mLabelString.Append("^A1N,").Append(vC_H).Append(",").Append(vC_W); //Scalable
                    mLabelString.Append("^FD");    //START DATA
                    mLabelString.Append(pManageF); //출력될 글자
                    mLabelString.Append("^FS");    //END DATA
                    //------------------------------------------------------------------------------------------------------------
                    // 관리자(부)
                    //------------------------------------------------------------------------------------------------------------
                    vObject = mDataTable.Rows[0]["D_FO_X"];
                    string vD_FO_X = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["D_FO_Y"];
                    string vD_FO_Y = ConvertNumberToString(vObject);
                    mLabelString.Append("^FO").Append(vD_FO_X).Append(",").Append(vD_FO_Y); //기준위치에서 지정한 위치에[열,행]

                    //mLabelString.Append("^A1N,36,20"); //Scalable
                    vObject = mDataTable.Rows[0]["D_H"];
                    string vD_H = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["D_W"];
                    string vD_W = ConvertNumberToString(vObject);
                    mLabelString.Append("^A1N,").Append(vD_H).Append(",").Append(vD_W); //Scalable
                    mLabelString.Append("^FD"); //START DATA
                    mLabelString.Append(pManageS); //출력될 글자
                    mLabelString.Append("^FS"); //END DATA
                    //------------------------------------------------------------------------------------------------------------
                    // 취득일자
                    //------------------------------------------------------------------------------------------------------------
                    vObject = mDataTable.Rows[0]["E_FO_X"];
                    string vE_FO_X = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["E_FO_Y"];
                    string vE_FO_Y = ConvertNumberToString(vObject);
                    mLabelString.Append("^FO").Append(vE_FO_X).Append(",").Append(vE_FO_Y); //기준위치에서 지정한 위치에[열,행]

                    //mLabelString.Append("^A1N,36,20"); //Scalable
                    vObject = mDataTable.Rows[0]["E_H"];
                    string vE_H = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["E_W"];
                    string vE_W = ConvertNumberToString(vObject);
                    mLabelString.Append("^A1N,").Append(vE_H).Append(",").Append(vE_W); //Scalable
                    mLabelString.Append("^FD"); //START DATA
                    mLabelString.Append(pAcquireDate); //출력될 글자
                    mLabelString.Append("^FS"); //END DATA
                    //------------------------------------------------------------------------------------------------------------
                    // 사용부서
                    //------------------------------------------------------------------------------------------------------------
                    vObject = mDataTable.Rows[0]["F_FO_X"];
                    string vF_FO_X = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["F_FO_Y"];
                    string vF_FO_Y = ConvertNumberToString(vObject);
                    mLabelString.Append("^FO").Append(vF_FO_X).Append(",").Append(vF_FO_Y); //기준위치에서 지정한 위치에[열,행]

                    //mLabelString.Append("^A1N,36,20"); //Scalable
                    vObject = mDataTable.Rows[0]["F_H"];
                    string vF_H = ConvertNumberToString(vObject);
                    vObject = mDataTable.Rows[0]["F_W"];
                    string vF_W = ConvertNumberToString(vObject);
                    mLabelString.Append("^A1N,").Append(vF_H).Append(",").Append(vF_W); //Scalable
                    mLabelString.Append("^FD"); //START DATA
                    mLabelString.Append(pUseDept); //출력될 글자
                    mLabelString.Append("^FS"); //END DATA
                    mLabelString.Append("^PQ1,1,1,Y^FS");
                    mLabelString.Append("^XZ"); //END
                    mLabelString.Append("\r\n");

                    /*
                    mLabelString.Append("^XA");
                    mLabelString.Append("^BY2,2.0^FS");
                    mLabelString.Append("^SEE:UHANGUL.DAT^FS");
                    mLabelString.Append("^CW1,E:KFONT3.FNT^CI26^FS");
                    mLabelString.Append("^FO50,50^A1N,40,40^FD한글출력테스트^FS");
                    mLabelString.Append("^PQ1,1,1,Y^FS");
                    mLabelString.Append("^XZ");
                    */

                    isMake = true;
                }
                catch (System.Exception ex)
                {
                    mMessageError = ex.Message;
                    mAppInterfaceAdv.OnAppMessage(mMessageError);
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            return isMake;
        }

        #endregion;

        #region ----- Printing Method -----

        public void Printing(object pManageNo, object pAssetName, object pItemSpec, object pManageF, object pManageS, object pAcquireDate, object pUseDept)
        {
            string vManageNo = ConvertString(pManageNo);  //관리번호
            string vAssetName = ConvertString(pAssetName);//자산명
            string vItemSpec = ConvertString(pItemSpec);  //규격
            string vManageF = ConvertString(pManageF);    //관리자(정)
            string vManageS = ConvertString(pManageS);    //관리자(부)
            //--------------------------------------------//취득일자
            System.DateTime vDateTime = new DateTime();   
            string vAcquireDate = string.Empty;            
            if (pAcquireDate != null)
            {
                vDateTime = (System.DateTime)pAcquireDate;
                vAcquireDate = vDateTime.ToString("yyyy년 MM월 dd일", null);
            }
            //----------------------------------------------------------------
            string vUseDept = ConvertString(pUseDept);    //사용부서

            bool isMake = MakeStringZPL(vManageNo, vAssetName, vItemSpec, vManageF, vManageS, vAcquireDate, vUseDept);

            string vLabelString = mLabelString.ToString();
            Label_Printing(vLabelString);

            //mAppInterfaceAdv.OnAppMessage(vLabelString);
            //System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Printing Method -----

        #region ----- Text File Export Methods ----

        private string ExportTXT(string pLabelString)
        {
            if (pLabelString == null)
            {
                return null;
            }

            System.IO.FileStream vWriteFile = null;
            System.Text.StringBuilder vSaveString = new System.Text.StringBuilder();

            string vSaveTextFileName = "FCMF0391_001.txt";
            string vPathZPL = string.Format("{0}\\{1}", System.Windows.Forms.Application.StartupPath, "Report");
            //string vPathZPL = @"D:\Project\FX_ERP\FX_Main\bin\Debug\Report";
            string vPathZPLtext = string.Format("{0}\\{1}", vPathZPL, vSaveTextFileName); 

            try
            {
                vWriteFile = System.IO.File.Open(vPathZPLtext, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None);

                vSaveString.Append(pLabelString);

                byte[] vSaveBytes = new System.Text.UTF8Encoding(true).GetBytes(vSaveString.ToString());
                int vSaveStrigLength = vSaveBytes.Length;

                vWriteFile.Write(vSaveBytes, 0, vSaveStrigLength);
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                mAppInterfaceAdv.OnAppMessage(vMessage);
            }
            mAppInterfaceAdv.OnAppMessage("Export Text End");

            vWriteFile.Dispose();

            return vSaveTextFileName;//vPathZPLtext;
        }
        #endregion;

        private void Label_Printing(string pLabelString)
        {
            try
            {
                Library_Parallel_PORT.LPT_PORT LabelPrint = new Library_Parallel_PORT.LPT_PORT();
                /*
                // pLabelString으로 받은 데이터를 txt 파일로 Export 시킨다.
                string sPathZPLtext = ExportTXT(pLabelString);f

                //LabelPrint.OutputPRiNTER(@"c:\LPT1:", pLabelString);
                //LabelPrint.OutputPRiNTER(@"c:\COM9:", pLabelString);

                string sStartFilePath = string.Format("/c COPY {0} LPT1:", sPathZPLtext); // /c는 Dos모드에서의 내부 명령임.
                System.Diagnostics.Process.Start("cmd.exe", sStartFilePath);
                //System.Diagnostics.Process.Start("cmd.exe", "/c COPY print.txt COM9:");
                */

                //string sPathZPLtext = ExportTXT(pLabelString);
                string sPathName = ExportTXT(pLabelString);
                string sPrintFileName = string.Format("Report\\{0}", sPathName);

                //string sStartFilePath = string.Format("/c COPY {0} LPT1:", sPrintFileName);
                string sStartFilePath = string.Format("/c COPY {0} COM5:", sPrintFileName);
                System.Diagnostics.ProcessStartInfo mProcessInfo = new System.Diagnostics.ProcessStartInfo("cmd.exe", sStartFilePath);
                mProcessInfo.CreateNoWindow = true;
                mProcessInfo.UseShellExecute = false;
                System.Diagnostics.Process mProcess = System.Diagnostics.Process.Start(mProcessInfo);
                mProcess.WaitForExit(500);
                mProcess.Close();

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterfaceAdv.OnAppMessage(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;

        #region ----- Method -----

        #endregion;
    }
}
