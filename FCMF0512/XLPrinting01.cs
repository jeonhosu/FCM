using System;

namespace FCMF0512
{
    public class XLPrinting01 : XLInterface
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        private int mPrintingLineFIRST = 10; //��� ���� �� ���� ��ġ

        private int mPrintingLineSTART = 10; //���� ��½� ���� ���� �� ��ġ ����

        private int mCopyLineSUM = 1;        //������ ���õ� ��Ʈ�� ����Ǿ��� ���� �� ��ġ
        private int mIncrementCopyMAX = 41;  //����Ǿ��� ���� ����

        private int mCopyColumnSTART = 1; //����Ǿ�  �� �� ���� ��
        private int mCopyColumnEND = 66;  //������ ���õ� ��Ʈ�� ����Ǿ��� �� �� ��ġ

        private string mtmpString1 = string.Empty;
        private string mtmpString2 = string.Empty;

        private string mMessageValue1 = string.Empty; //�Ұ�[EAPP_10046]
        private string mMessageValue2 = string.Empty; //�Ѱ����ڱ�[FCM_10222]
        private string mMessageValue3 = string.Empty; //�ѱ�����ä

        private bool mIsPrinted = false; //�ڱ���Ȳ�� ��� �ߴ���?

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

        public int PrintingLineSTART
        {
            set
            {
                mPrintingLineSTART = value;
            }
            get
            {
                return mPrintingLineSTART;
            }
        }

        public int CopyLineSUM
        {
            set
            {
                mCopyLineSUM = value;
            }
            get
            {
                return mCopyLineSUM;
            }
        }

        public int PrintingLineFIRST
        {
            set
            {
                mPrintingLineFIRST = value;
            }
            get
            {
                return mPrintingLineFIRST;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLPrinting01(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterface;
            mMessageAdapter = pMessageAdapter;
        }

        public XLPrinting01(XL.XLPrint pPrinting, InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mPrinting = pPrinting;
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

        #region ----- Convert DateTime Methods ----

        private object ConvertDateTime(object pObject)
        {
            object vObject = null;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd HH:mm:ss", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vObject;
        }

        private object ConvertDate(object pObject)
        {
            object vObject = null;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vObject;
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

        #region ----- Line Clear All Methods ----

        private void XlLineClear(int pPrintingLine)
        {
            try
            {
                mPrinting.XLActiveSheet("Destination");

                int vStartRow = pPrintingLine + 1;
                int vStartCol = mCopyColumnSTART + 1;
                int vEndRow = mCopyLineSUM - 3;
                int vEndCol = mCopyColumnEND - 1;

                if (pPrintingLine > vEndRow)
                {
                    return;
                }

                if (vStartRow > vEndRow)
                {
                    vStartRow = vEndRow; //�����ϴ� ���� �����, ������ �� ���� ���� Ŀ���Ƿ�, ������ �� ���� ��
                }

                if (vStartRow == vEndRow)
                {
                    mPrinting.XL_LineClear(vStartRow, vStartCol, vEndCol);
                }
                else
                {
                    mPrinting.XL_LineClearInSide(vStartRow, vStartCol, vEndRow, vEndCol);
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }
        }

        #endregion;

        #region ----- Cell Merge Methods ----

        private void CellMerge(int pXLine, int[] pXLColumn, string pMessageValue)
        {
            try
            {
                int vXLine = pXLine - 1;

                mPrinting.XLActiveSheet("Destination");

                object vObject = null;
                mPrinting.XLSetCell(vXLine, pXLColumn[0], vObject);
                mPrinting.XLSetCell(vXLine, pXLColumn[1], vObject);
                mPrinting.XLSetCell(vXLine, pXLColumn[2], vObject);

                int vStartRow = vXLine;
                int vStartCol = pXLColumn[0];
                int vEndRow = vXLine;
                int vEndCol = pXLColumn[3] - 1;

                mPrinting.XLCellMerge(vStartRow, vStartCol, vEndRow, vEndCol, false);

                mPrinting.XLSetCell(vXLine, pXLColumn[0], pMessageValue);

                mPrinting.XL_LineDraw_TopBottom(vXLine, vStartCol, (mCopyColumnEND - 1), 2);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }
        }

        #endregion;

        #region ----- Line SLIP Methods ----

        #region ----- Array Set ----

        private void SetArray(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[10];
            pXLColumn = new int[10];

            pGDColumn[0] = pGrid.GetColumnToIndex("TR_CLASS_NAME");        //����
            pGDColumn[1] = pGrid.GetColumnToIndex("TR_MANAGE_NAME");       //����
            pGDColumn[2] = pGrid.GetColumnToIndex("BANK_NAME");            //�����
            pGDColumn[3] = pGrid.GetColumnToIndex("BEGIN_AMOUNT");         //�����ܾ�
            pGDColumn[4] = pGrid.GetColumnToIndex("DR_AMOUNT");            //�����Ա�
            pGDColumn[5] = pGrid.GetColumnToIndex("CR_AMOUNT");            //�������
            pGDColumn[6] = pGrid.GetColumnToIndex("REMAIN_AMOUNT");        //�����ܾ�
            pGDColumn[7] = pGrid.GetColumnToIndex("CURRENCY_CODE");        //��ȭ
            pGDColumn[8] = pGrid.GetColumnToIndex("REMAIN_CURR_AMOUNT");   //�����ܾ�[��ȭ�ܾ�]
            pGDColumn[9] = pGrid.GetColumnToIndex("DESCRIPTION");          //���

            pXLColumn[0] = 2;    //����
            pXLColumn[1] = 6;    //����
            pXLColumn[2] = 10;   //�����
            pXLColumn[3] = 19;   //�����ܾ�
            pXLColumn[4] = 27;   //�����Ա�
            pXLColumn[5] = 35;   //�������
            pXLColumn[6] = 43;   //�����ܾ�
            pXLColumn[7] = 52;   //��ȭ
            pXLColumn[8] = 54;   //�����ܾ�[��ȭ�ܾ�]
            pXLColumn[9] = 59;   //���
        }

        #endregion;

        #region ----- Convert decimal  Method ----

        private decimal ConvertNumber(string pStringNumber)
        {
            decimal vConvertDecimal = 0m;

            try
            {
                bool isNull = string.IsNullOrEmpty(pStringNumber);
                if (isNull != true)
                {
                    vConvertDecimal = decimal.Parse(pStringNumber);
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vConvertDecimal;
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

        private bool IsConvertDate(object pObject, out string pConvertDateTimeShort)
        {
            bool vIsConvert = false;
            pConvertDateTimeShort = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        pConvertDateTimeShort = vDateTime.ToShortDateString();
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

        #region ----- XLHeader Methods -----

        private void XLHeader(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            string vTitleName = string.Empty;
            string vHeaderColumn_0 = string.Empty;
            string vHeaderColumn_1 = string.Empty;
            string vHeaderColumn_2 = string.Empty;

            int vXLineTitel = pXLine - 3; 
            int vXLine = pXLine - 1;

            vTitleName = "2. ����ä��";
            vHeaderColumn_0 = "���Ա���";
            vHeaderColumn_1 = pGrid.GridAdvExColElement[pGDColumn[4]].HeaderElement[0].TL1_KR;
            vHeaderColumn_2 = pGrid.GridAdvExColElement[pGDColumn[5]].HeaderElement[0].TL1_KR;

            mPrinting.XLSetCell(vXLineTitel, pXLColumn[0], vTitleName);
            mPrinting.XLSetCell(vXLine, pXLColumn[2], vHeaderColumn_0);
            mPrinting.XLSetCell(vXLine, pXLColumn[4], vHeaderColumn_1);
            mPrinting.XLSetCell(vXLine, pXLColumn[5], vHeaderColumn_2);
        }

        #endregion;

        #region ----- XlLine Methods -----

        private int XlLine(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
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
                mPrinting.XLActiveSheet("Destination");

                //[TR_MANAGE_NAME]����
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    if (mtmpString2 != vConvertString)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                        mtmpString2 = vConvertString;

                        int vDrawLineColumnSTART = vXLColumnIndex;
                        int vDrawLineColumnEND = mCopyColumnEND - 1;

                        mPrinting.XL_LineDraw_Top(vXLine, vDrawLineColumnSTART, (pXLColumn[2] - 1), 1);

                        if (mMessageValue1 == vConvertString)
                        {
                            vDrawLineColumnSTART = pXLColumn[1];
                            vDrawLineColumnEND = pXLColumn[3] - 1;
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex, null);
                            mPrinting.XLCellMerge(vXLine, vDrawLineColumnSTART, vXLine, vDrawLineColumnEND, false);
                            mPrinting.XLSetCell(vXLine, vDrawLineColumnSTART, vConvertString);

                            vDrawLineColumnSTART = pXLColumn[1];
                            vDrawLineColumnEND = mCopyColumnEND - 1;
                            mPrinting.XL_LineDraw_TopBottom(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND, 2);
                        }
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[TR_CLASS_NAME]����
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    if (mtmpString1 != vConvertString)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                        mtmpString1 = vConvertString;

                        int vDrawLineColumnSTART = mCopyColumnSTART + 1;
                        int vDrawLineColumnEND = mCopyColumnEND - 1;

                        mPrinting.XL_LineDraw_Top(vXLine, vDrawLineColumnSTART, vDrawLineColumnEND, 2);
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[BANK_NAME]����
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
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

                //[BEGIN_AMOUNT]�����ܾ�
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[DR_AMOUNT]�����Ա�
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[CR_AMOUNT]�������
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[REMAIN_AMOUNT]�����ܾ�
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[CURRENCY_CODE]��ȭ
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
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

                //[REMAIN_CURR_AMOUNT]�����ܾ�[��ȭ�ܾ�]
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0.00}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //[DESCRIPTION]���
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
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

                //-------------------------------------------------------------------
                vXLine++;
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

        public int LineWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx[] pGrid, string pDate)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;
            mIsPrinted = false;

            int[] vGDColumn;
            int[] vXLColumn;

            int vPrintingLine = mPrintingLineSTART;

            try
            {
                SetArray(pGrid[0], out vGDColumn, out vXLColumn);

                mPrinting.XLActiveSheet("SourceTab1");
                mPrinting.XLSetCell(4, 2, pDate);

                int vTotalRow = pGrid[0].RowCount;

                int vTotal = pGrid[0].RowCount + pGrid[1].RowCount;
                mPageTotalNumber = vTotal / 31;
                mPageTotalNumber = (vTotal % 31) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);

                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    mMessageValue1 = mMessageAdapter.ReturnText("EAPP_10046");  //�Ұ�[EAPP_10046]
                    mMessageValue2 = mMessageAdapter.ReturnText("FCM_10222");   //�Ѱ����ڱ�[FCM_10222]

                    if (vTotalRow > 31)
                    {
                        mIncrementCopyMAX = 41;
                    }
                    else
                    {
                        mIncrementCopyMAX = vTotalRow + 9;
                    }
                    mCopyLineSUM = FirstCopyAndPaste(mPrinting, mCopyLineSUM);

                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vPrintingLine = XlLine(pGrid[0], vRow, vPrintingLine, vGDColumn, vXLColumn);
                        mIsPrinted = true;

                        if (vTotalRow == vCountRow)
                        {
                            CellMerge(vPrintingLine, vXLColumn, mMessageValue2);
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);
                                mtmpString1 = string.Empty;
                                mtmpString2 = string.Empty;
                            }
                        }
                    }
                }

                vTotalRow = pGrid[1].RowCount;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    mMessageValue3 = "�ѱ�����ä";

                    if (mIsPrinted == false)
                    {
                        if (vTotalRow > 31)
                        {
                            mIncrementCopyMAX = 41;
                        }
                        else
                        {
                            mIncrementCopyMAX = vTotalRow + 9;
                        }

                        //�ڱ���Ȳ�� ��� �� �ߴٸ�
                        mCopyLineSUM = FirstCopyAndPaste(mPrinting, mCopyLineSUM);

                        vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);
                    }
                    else
                    {
                        mCopyLineSUM++;

                        if (vTotalRow > 31)
                        {
                            mIncrementCopyMAX = 41;
                        }
                        else
                        {
                            mIncrementCopyMAX = vTotalRow + 9;
                        }

                        //�ڱ���Ȳ�� ��� �ߴٸ�
                        mCopyLineSUM = SecondCopyAndPaste(mPrinting, mCopyLineSUM);

                        vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);
                    }

                    XLHeader(pGrid[1], vPrintingLine, vGDColumn, vXLColumn);

                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vPrintingLine = XlLine(pGrid[1], vRow, vPrintingLine, vGDColumn, vXLColumn);

                        if (vTotalRow == vCountRow)
                        {
                            CellMerge(vPrintingLine, vXLColumn, mMessageValue3);
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineFIRST - 1);
                                mtmpString1 = string.Empty;
                                mtmpString2 = string.Empty;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            mPrintingLineSTART = vPrintingLine;

            return mPageNumber;
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine)
        {
            int vPrintingLineEND = mCopyLineSUM - 1;
            if (vPrintingLineEND < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = SecondCopyAndPaste(mPrinting, mCopyLineSUM);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //ù��° ������ ����
        private int FirstCopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            pPrinting.XLActiveSheet("SourceTab1");
            object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet("Destination");
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);


            if (mPageTotalNumber > 1)
            {
                mPageNumber++; //������ ��ȣ
                mPrinting.XLCellMerge(vCopyPrintingRowEnd, 33, vCopyPrintingRowEnd, 36, false);
                string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
                mPrinting.XLSetCell((vCopyPrintingRowEnd - 1), 33, vPageNumberText); //������ ��ȣ, XLcell[��, ��]
            }

            return vCopySumPrintingLine;
        }

        //�ι�° ������ ����
        private int SecondCopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPageNumber++; //������ ��ȣ

            int vIncrementCopyMAX = mIncrementCopyMAX - 6;
            int vCopyColumnSTART = mCopyColumnSTART + 6;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + vIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            pPrinting.XLActiveSheet("SourceTab1");
            object vRangeSource = pPrinting.XLGetRange(vCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet("Destination");
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            mPrinting.XLCellMerge(vCopyPrintingRowEnd, 33, vCopyPrintingRowEnd, 36, false);

            string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
            mPrinting.XLSetCell(vCopyPrintingRowEnd, 33, vPageNumberText); //������ ��ȣ, XLcell[��, ��]

            mPrinting.XL_LineClearTOP(vCopyPrintingRowSTART, mCopyColumnSTART, mCopyColumnEND);

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND);
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xlsx", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;
    }
}