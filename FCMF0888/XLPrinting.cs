using System;
using ISCommonUtil;

namespace FCMF0888
{
    public class XLPrinting
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private int mPageNumber = 0;

        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        private int mPrintingLineSTART_SOURCE1 = 10;  //�����ſ��� ����Ȯ�μ��� ���� ���޽��� �հ�
        private int mPrintingLineSTART_SOURCE2 = 8;   //�����ſ��� ����Ȯ�μ��� ���� ���޽��� ����

        private int mCopyLineSUM = 1;        //������ ���õ� ��Ʈ�� ����Ǿ��� ���� �� ��ġ, ���� �� ����
        private int mIncrementCopyMAX = 38;  //����Ǿ��� ���� ����

        private int mCopyColumnSTART = 1; //����Ǿ�  �� �� ���� ��
        private int mCopyColumnEND = 44;  //������ ���õ� ��Ʈ�� ����Ǿ��� �� �� ��ġ

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        public string OpenFileNameExcel
        {
            set
            {
                mXLOpenFileName = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterface;
            mMessageAdapter = pMessageAdapter;
        }

        #endregion;

        #region ----- XL File Open -----

        public bool XLFileOpen()
        {
            bool IsOpen = false;

            try
            {
                IsOpen = mPrinting.XLOpenFile(mXLOpenFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }

            return IsOpen;
        }

        #endregion;

        #region ----- Dispose -----

        public void Dispose()
        {
            mPrinting.XLOpenFileClose();
            mPrinting.XLClose();
        }

        #endregion;

        #region ----- MaxIncrement Methods ----

        private int MaxIncrement(string pPathBase, string pSaveFileName)
        {
            int vMaxNumber = 0;
            System.IO.DirectoryInfo vFolder = new System.IO.DirectoryInfo(pPathBase);
            string vPattern = string.Format("{0}*", pSaveFileName);
            System.IO.FileInfo[] vFiles = vFolder.GetFiles(vPattern);

            foreach (System.IO.FileInfo vFile in vFiles)
            {
                string vFileNameExt = vFile.Name;
                int vCutStart = vFileNameExt.LastIndexOf(".");
                string vFileName = vFileNameExt.Substring(0, vCutStart);

                int vCutRight = 3;
                int vSkip = vFileName.Length - vCutRight;
                string vTextNumber = vFileName.Substring(vSkip, vCutRight);
                int vNumber = int.Parse(vTextNumber);

                if (vNumber > vMaxNumber)
                {
                    vMaxNumber = vNumber;
                }
            }

            return vMaxNumber;
        }

        #endregion;

        #region ----- Line SLIP Methods ----

        #region ----- Array Set 1 ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[5];
            pXLColumn = new int[5];

            pGDColumn[0] = pGrid.GetColumnToIndex("VAT_RECEIPT_DESC");  //����
            pGDColumn[1] = pGrid.GetColumnToIndex("SUPPLIER_COUNT");    //����ó��
            pGDColumn[2] = pGrid.GetColumnToIndex("VAT_COUNT");         //�Ǽ�
            pGDColumn[3] = pGrid.GetColumnToIndex("ITEM_AMOUNT");       //�ݾ�(��)
            pGDColumn[4] = pGrid.GetColumnToIndex("DEEMED_VAT_AMOUNT"); //�������Լ���

            pXLColumn[0] = 1;   //����
            pXLColumn[1] = 13;  //����ó ��
            pXLColumn[2] = 21;  //�Ǽ�
            pXLColumn[3] = 27;  //���ݾ�
            pXLColumn[4] = 35;  //���Լ��� ������
        }

        #endregion;

        #region ----- Array Set 2 ----

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[6];
            pXLColumn = new int[6];

            pGDColumn[0] = pGrid.GetColumnToIndex("SEQ");        //��ȣ
            pGDColumn[1] = pGrid.GetColumnToIndex("VAT_DOMESTIC_LC_NM"); //���и�
            pGDColumn[2] = pGrid.GetColumnToIndex("DOC_NO");     //������ȣ
            pGDColumn[3] = pGrid.GetColumnToIndex("PUB_DATE");   //�߱���
            pGDColumn[4] = pGrid.GetColumnToIndex("VAT_NUMBER"); //����ڵ�Ϲ�ȣ
            pGDColumn[5] = pGrid.GetColumnToIndex("SUPPLY_AMT"); //�ݾ�(��)


            pXLColumn[0] = 2;   //��ȣ
            pXLColumn[1] = 4;   //���и�
            pXLColumn[2] = 10;  //������ȣ
            pXLColumn[3] = 18;  //�߱���
            pXLColumn[4] = 23;  //����ڵ�Ϲ�ȣ
            pXLColumn[5] = 31;  //�ݾ�(��)
        }

        #endregion;

        #region ----- IsConvert Methods -----

        private bool IsConvertString(object pObject, out string pConvertString)
        {
            bool vIsConvert = false;
            pConvertString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is string;
                    if (vIsConvert == true)
                    {
                        pConvertString = pObject as string;
                    }
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        private bool IsConvertNumber(object pObject, out decimal pConvertDecimal)
        {
            bool vIsConvert = false;
            pConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is decimal;
                    if (vIsConvert == true)
                    {
                        decimal vIsConvertNum = (decimal)pObject;
                        pConvertDecimal = vIsConvertNum;
                    }
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        private bool IsConvertDate(object pObject, out System.DateTime pConvertDateTimeShort)
        {
            bool vIsConvert = false;
            pConvertDateTimeShort = new System.DateTime();

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is System.DateTime;
                    if (vIsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        pConvertDateTimeShort = vDateTime;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        #endregion;

        #region ----- Header Write Method ----

        private void XLHeader(InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PRINT_COPPER_ETC_TITLE)
        {
            int vCountRow = 0;
            int vXLine = 0;
            int vXLColumn = 0;
            object vObject = null;
            string vConvertString = string.Empty;
            bool IsConvert = false;

            try
            {
                vCountRow = pIDA_PRINT_COPPER_ETC_TITLE.CurrentRows.Count;

                if (vCountRow > 0)
                {
                    mPrinting.XLActiveSheet("SOURCE2");

                    //�ΰ���ġ���Ű���
                    vObject = pIDA_PRINT_COPPER_ETC_TITLE.OraSelectData.Rows[0]["FISCAL_YEAR"];
                    vXLine = 3;
                    vXLColumn = 6;
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }

                    //����ڵ�Ϲ�ȣ
                    vObject = pIDA_PRINT_COPPER_ETC_TITLE.OraSelectData.Rows[0]["VAT_NUMBER"];
                    vXLine = 4;
                    vXLColumn = 32;
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }


                    mPrinting.XLActiveSheet("SOURCE1");

                    //�ΰ���ġ���Ű���
                    vObject = pIDA_PRINT_COPPER_ETC_TITLE.OraSelectData.Rows[0]["FISCAL_YEAR"];
                    vXLine = 3;
                    vXLColumn = 6;
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }

                    //��ȣ(���θ�)
                    vObject = pIDA_PRINT_COPPER_ETC_TITLE.OraSelectData.Rows[0]["CORP_NAME"];
                    vXLine = 6;
                    vXLColumn = 8;
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }

                    //����ڵ�Ϲ�ȣ
                    vObject = pIDA_PRINT_COPPER_ETC_TITLE.OraSelectData.Rows[0]["VAT_NUMBER"];
                    vXLine = 6;
                    vXLColumn = 29;
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }

                    //����
                    vObject = pIDA_PRINT_COPPER_ETC_TITLE.OraSelectData.Rows[0]["BUSINESS_ITEM"];
                    vXLine = 7;
                    vXLColumn = 8;
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }

                    //����
                    vObject = pIDA_PRINT_COPPER_ETC_TITLE.OraSelectData.Rows[0]["BUSINESS_TYPE"];
                    vXLine = 7;
                    vXLColumn = 29;
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }

                    //�ۼ�����
                    vObject = pIDA_PRINT_COPPER_ETC_TITLE.OraSelectData.Rows[0]["CREATE_DATE"];
                    vXLine = 33;
                    vXLColumn = 27;
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }

                    //������
                    vObject = pIDA_PRINT_COPPER_ETC_TITLE.OraSelectData.Rows[0]["REPORTED_BY"];
                    vXLine = 34;
                    vXLColumn = 22;
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }

                    //���Ҽ�����
                    vObject = pIDA_PRINT_COPPER_ETC_TITLE.OraSelectData.Rows[0]["TAX_OFFICE_NAME"];
                    vXLine = 35;
                    vXLColumn = 1;
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- 3. ��Ȱ�����ڿ� ���Լ��װ��� ���� �Ű��� Write Method -----

        private void XLLine_Report(System.Data.DataRow pRow)
        {
            int vXLine = 17; //������ ������ ǥ�õǴ� �� ��ȣ

            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("DESTINATION");

                //����� �հ� 													 
                vXLColumnIndex = 1;
                vObject = pRow["SALES_SUM_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����� ������
                vXLColumnIndex = 6;
                vObject = pRow["SALES_PRE_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����� Ȯ����
                vXLColumnIndex = 11;
                vObject = pRow["SALES_FIX_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�ѵ���
                vXLColumnIndex = 16;
                vObject = string.Format("{0}/{1}", pRow["LIMIT_RATE_NUMERATOR"], pRow["LIMIT_RATE_DENOMINATOR"]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�ѵ���
                vXLColumnIndex = 20;
                vObject = pRow["LIMIT_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�հ�
                vXLColumnIndex = 25;
                vObject = pRow["PURCHASES_SUM_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�����Ծ� - ���ݰ�꼭 
                vXLColumnIndex = 30;
                vObject = pRow["PURCHASES_TAX_BILL_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�����Ծ�-��������
                vXLColumnIndex = 35;
                vObject = pRow["PURCHASES_BILL_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�������ɱݾ�
                vXLColumnIndex = 40;
                vObject = pRow["DED_RANGE_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------

                //�������ݾ�
                vXLColumnIndex = 1;
                vObject = pRow["DED_TARGET_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //������󼼾� - ������ 
                vXLColumnIndex = 11;
                vObject = string.Format("{0}/{1}", pRow["DED_RATE_NUMERATOR"], pRow["DED_RATE_DENOMINATOR"]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //������󼼾� - ������� ���� 
                vXLColumnIndex = 15;
                vObject = pRow["DED_VAT_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�̹̰�����������-�հ� 
                vXLColumnIndex = 22;
                vObject = pRow["DED_PRE_VAT_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�̹̰�����������-�����Ű��
                vXLColumnIndex = 27;
                vObject = pRow["DED_PRE_QUARTER_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�̹̰�����������-��������Ű��
                vXLColumnIndex = 32;
                vObject = pRow["DED_PRE_MONTHLY_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����(����)�� ����
                vXLColumnIndex = 37;
                vObject = pRow["FIX_VAT_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal < 0)
                    {
                        vConvertString = string.Format("��{0:##,###,###,###,###,###,###,###,###}", Math.Abs(vConvertDecimal));
                    }
                    else
                    {
                        vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Line 1 Write Method -----

        private int XLLine_1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("DESTINATION");
                //����
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����ó��
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                //�Ǽ�
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //���ݾ�
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�������Լ���
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }


            pXLine = vXLine;

            return pXLine;
        }

        #endregion;

        #region ----- Line 2 Write Method -----

        private int XLLine_2(System.Data.DataRow pRow, int pXLine)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("DESTINATION");

                //��ȣ
                vXLColumnIndex = 1;
                vObject = pRow["SEQ"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //���� �Ǵ� ��ȣ
                vXLColumnIndex = 3;
                vObject = pRow["SUPPLIER_NAME"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�ֹε�Ϲ�ȣ �Ǵ� ����ڵ�Ϲ�ȣ 
                vXLColumnIndex = 13;
                vObject = pRow["TAX_REG_NO"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�Ǽ�
                vXLColumnIndex = 19;
                vObject = pRow["VAT_COUNT"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //ǰ��
                vXLColumnIndex = 21;
                vObject = pRow["ITEM_DESC"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����
                vXLColumnIndex = 28;
                vObject = pRow["ITEM_QTY"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�����ȣ
                vXLColumnIndex = 30;
                vObject = pRow["CAR_NUM"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�����ȣ
                vXLColumnIndex = 34;
                vObject = pRow["CAR_BODY_NUM"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //���ݾ�
                vXLColumnIndex = 40;
                vObject = pRow["ITEM_AMOUNT"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }


            pXLine = vXLine;

            return pXLine;
        }

        #endregion;

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pIGR_SUM_RECYCLING_ETC
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_RECYCLING_ETC_REPORT
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PRINT_RECYCLING_ETC
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PRINT_RECYCLING_ETC_TITLE)
        {
            mPageNumber = 0;
            string vMessage = string.Empty;

            int[] vGDColumn;
            int[] vXLColumn;

            int vPrintingLine = 0;
            int vCountROW1 = 0;

            try
            {
                int vTotal1ROW = pIGR_SUM_RECYCLING_ETC.RowCount;
                int vTotal2ROW = pIDA_RECYCLING_ETC_REPORT.CurrentRows.Count;
                int vTotal3ROW = pIDA_PRINT_RECYCLING_ETC.CurrentRows.Count;

                #region ----- Header/SUM Write ----

                XLHeader(pIDA_PRINT_RECYCLING_ETC_TITLE);

                mCopyLineSUM = CopyAndPaste("SOURCE1", mCopyLineSUM);

                #endregion;

                #region ----- Line 1 Write ----

                if (vTotal1ROW > 0)
                {
                    vCountROW1 = 0;

                    SetArray1(pIGR_SUM_RECYCLING_ETC, out vGDColumn, out vXLColumn);

                    vPrintingLine = mPrintingLineSTART_SOURCE1;

                    for (int vRow1 = 0; vRow1 < vTotal1ROW; vRow1++)
                    {
                        vCountROW1++;

                        vMessage = string.Format("Grid1_Rows : {0}/{1}", vCountROW1, vTotal1ROW);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vPrintingLine = XLLine_1(pIGR_SUM_RECYCLING_ETC, vRow1, vPrintingLine, vGDColumn, vXLColumn);
                    }
                }

                #endregion;

                #region ----- 3. ��Ȱ�����ڿ� ���Լ��װ��� ���� �Ű��� Line 2 Write ----

                if (vTotal2ROW > 0)
                {
                    vCountROW1 = 0;

                    foreach (System.Data.DataRow vRow in pIDA_RECYCLING_ETC_REPORT.CurrentRows)
                    {
                        vCountROW1 = vCountROW1 + 1;
                        vMessage = string.Format("Print Data Line : {0}/{1}", vCountROW1, vTotal3ROW);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        XLLine_Report(vRow);

                        if (vTotal2ROW == vCountROW1)
                        {
                            //������ ���̸�...
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART_SOURCE2 - 1);
                            }
                        }
                    }
                }
                #endregion;

                #region ----- Line 3 Write ----

                if (vTotal3ROW > 0)
                {
                    vCountROW1 = 0;

                    vPrintingLine = vPrintingLine + 13; // 17 = 13 + 4
                    foreach (System.Data.DataRow vRow in pIDA_PRINT_RECYCLING_ETC.CurrentRows)
                    {
                        vCountROW1 = vCountROW1 + 1;
                        vMessage = string.Format("Print Data Line : {0}/{1}", vCountROW1, vTotal3ROW);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vPrintingLine = XLLine_2(vRow, vPrintingLine);

                        if (vTotal3ROW == vCountROW1)
                        {
                            //������ ���̸�...
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART_SOURCE2 - 1);
                            }
                        }
                    }
                }

                #endregion;
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine)
        {
            int vPrintingLineEND = mCopyLineSUM - 8; //1~37: mCopyLineSUM=38���� ������ ��µǴ� ���� 31 ���� �̹Ƿ�, 7�� ���� �ȴ�
            if (mPageNumber != 1)
            {
                //ù �������� ��� ������
                if (vPrintingLineEND < pPrintingLine)
                {
                    mIsNewPage = true;
                    mCopyLineSUM = CopyAndPaste("SOURCE2", mCopyLineSUM);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
            else
            {
                //ù��° ������ ���
                if (30 < pPrintingLine)
                {
                    mIsNewPage = true;
                    mCopyLineSUM = CopyAndPaste("SOURCE2", mCopyLineSUM);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //ù��° ������ ����
        private int CopyAndPaste(string pSheetTab, int pCopySumPrintingLine)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;

            mPrinting.XLActiveSheet(pSheetTab);

            object vRangeSource = mPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            mPrinting.XLActiveSheet("DESTINATION");
            object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            mPageNumber++; //������ ��ȣ
            int vRowPRINT = vCopyPrintingRowEnd - 3; // 35 = 38 - 3
            int vColumnPRINT = 39;
            if (mPageNumber != 1)
            {
                mPrinting.XLSetCell(vRowPRINT, vColumnPRINT, mPageNumber); //������ ��ȣ, XLcell[��, ��]
            }

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;
    }
}
