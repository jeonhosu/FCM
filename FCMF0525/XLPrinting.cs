using System;

namespace FCMF0525
{
    public class XLPrinting
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private int mPageTotalNumber = 0;
        private int mPageNumber = 0;
        
        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        private int mPrintingLineSTART1 = 10; //라인 출력시 엑셀 시작 행 위치 지정
        private int mPrintingLineEND1 = 38;   //mPrintingLineSTART1 부터 실제 출력될 마지막 행 위치 지정

        private int mPrintingLineSTART2 = 5;  //라인 출력시 엑셀 시작 행 위치 지정
        private int mPrintingLineEND2 = 58;   //mPrintingLineSTART2 부터 실제 출력될 마지막 행 위치 지정

        private int mCopyLineSUM = 1;        //엑셀의 선택된 쉬트의 복사되어질 시작 행 위치
        private int mIncrementCopyMAX = 41;  //복사되어질 행의 범위

        private int mCopyColumnSTART = 1; //복사되어진 행 누적 수
        private int mCopyColumnEND = 69;  //엑셀의 선택된 쉬트의 복사되어질 끝 열 위치

        private decimal mDR_AMOUNT = 0; //차변합계
        private decimal mCR_AMOUNT = 0; //대변합계

        private bool mFirstPagePrinted = false;

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

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterface;
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

                int vCutRight = 2;
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

        #region ----- Content Clear All Methods ----

        private void XlAllContentClear(XL.XLPrint pPrinting)
        {
            object vObject = null;

            if (mFirstPagePrinted == false)
            {
                pPrinting.XLActiveSheet("SourceTab1");

                int vStartRow = mPrintingLineSTART1;
                int vStartCol = mCopyColumnSTART + 1;
                int vEndRow = mPrintingLineEND1;
                int vEndCol = mCopyColumnEND - 1;

                mPrinting.XLSetCell(vStartRow, vStartCol, vEndRow, vEndCol, vObject);
            }
            else
            {
                pPrinting.XLActiveSheet("SourceTab2");

                int vStartRow = mPrintingLineSTART2;
                int vStartCol = mCopyColumnSTART + 1;
                int vEndRow = mPrintingLineEND2;
                int vEndCol = mCopyColumnEND - 1;

                mPrinting.XLSetCell(vStartRow, vStartCol, vEndRow, vEndCol, vObject);
            }
        }

        #endregion;

        #region ----- Line Clear All Methods ----

        private void XlLineClear(int pPrintingLine)
        {
            if (mFirstPagePrinted == false)
            {
                mPrinting.XLActiveSheet("SourceTab1");

                int vStartRow = pPrintingLine + 2;
                int vStartCol = mCopyColumnSTART + 1;
                int vEndRow = mPrintingLineEND1 - 1;
                int vEndCol = mCopyColumnEND - 1;

                if (vStartRow > vEndRow)
                {
                    vStartRow = vEndRow; //시작하는 행이 계산후, 끝나는 행 보다 값이 커지므로, 끝나는 행 값을 줌
                }

                mPrinting.XL_LineClearInSide(vStartRow, vStartCol, vEndRow, vEndCol);
            }
            else
            {
                mPrinting.XLActiveSheet("SourceTab2");

                int vStartRow = pPrintingLine + 2;
                int vStartCol = mCopyColumnSTART + 1;
                int vEndRow = mPrintingLineEND2 - 1;
                int vEndCol = mCopyColumnEND - 1;

                if (vStartRow > vEndRow)
                {
                    vStartRow = vEndRow;
                }

                mPrinting.XL_LineClearInSide(vStartRow, vStartCol, vEndRow, vEndCol);
            }
        }

        #endregion;

        #region ----- Excel Wirte [Header] Methods ----

        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter)
        {
            object vObject = null;

            try
            {
                //System.Drawing.Point vCellPoint01 = new System.Drawing.Point(7, 2);   //전표번호[GL_NUM]
                //System.Drawing.Point vCellPoint02 = new System.Drawing.Point(1, 31);  //발의일자[SLIP_DATE]
                //System.Drawing.Point vCellPoint03 = new System.Drawing.Point(3, 31);  //발의부서명[DEPT_NAME]
                //System.Drawing.Point vCellPoint04 = new System.Drawing.Point(5, 31);  //발의자 이름[PERSON_NAME]
                System.Drawing.Point vCellPoint05 = new System.Drawing.Point(6, 24);  //전표일자[GL_DATE]
                //System.Drawing.Point vCellPoint06 = new System.Drawing.Point(9, 2);   //적요[REMARK]

                mPrinting.XLActiveSheet("SourceTab1"); //셀에 문자를 넣기 위해 쉬트 선택

                ////전표번호[GL_NUM]
                //vObject = pAdapter.CurrentRow["GL_NUM"];
                //if (vObject != null)
                //{
                //    vObject = string.Format("전표번호 : {0}", vObject);
                //    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vObject);
                //}
                //else
                //{
                //    vObject = null;
                //    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vObject);
                //}

                ////작성일자[SLIP_DATE]
                //vObject = pAdapter.CurrentRow["BATCH_DATE"];
                //if (vObject != null)
                //{
                //    vObject = ConvertDate(vObject);
                //    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, vObject);
                //}
                //else
                //{
                //    vObject = null;
                //    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, vObject);
                //}

                ////작성부서명[DEPT_NAME]
                //vObject = pAdapter.CurrentRow["DEPT_NAME"];
                //if (vObject != null)
                //{
                //    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, vObject);
                //}
                //else
                //{
                //    vObject = null;
                //    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, vObject);
                //}

                ////작성자 이름[PERSON_NAME]
                //vObject = pAdapter.CurrentRow["PERSON_NAME"];
                //if (vObject != null)
                //{
                //    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, vObject);
                //}
                //else
                //{
                //    vObject = null;
                //    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, vObject);
                //}

                //전표일자[GL_DATE]
                vObject = pAdapter.CurrentRow["GL_DATE"];
                if (vObject != null)
                {
                    vObject = ConvertDate(vObject);
                    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, vObject);
                }

                ////적요[REMARK]
                //string vText = string.Empty;
                //vObject = pAdapter.CurrentRow["REMARK"];
                //if (vObject != null)
                //{
                //    bool isConvert = vObject is string;
                //    if (isConvert == true)
                //    {
                //        vText = vObject as string;
                //        bool isNull = string.IsNullOrEmpty(vText.Trim());
                //        if (isNull != true)
                //        {
                //            vText = string.Format("내역 : {0}", vObject);
                //        }
                //    }
                //    vObject = vText;
                //    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, vObject);
                //}
                //else
                //{
                //    vObject = null;
                //    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, vObject);
                //}
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        public void HeaderWrite_1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow)
        {
            object vObject = null;

            try
            {
                //System.Drawing.Point vCellPoint01 = new System.Drawing.Point(7, 2);   //전표번호[GL_NUM]
                //System.Drawing.Point vCellPoint02 = new System.Drawing.Point(1, 31);  //발의일자[SLIP_DATE]
                //System.Drawing.Point vCellPoint03 = new System.Drawing.Point(3, 31);  //발의부서명[DEPT_NAME]
                //System.Drawing.Point vCellPoint04 = new System.Drawing.Point(5, 31);  //발의자 이름[PERSON_NAME]
                System.Drawing.Point vCellPoint05 = new System.Drawing.Point(6, 24);  //전표일자[SLIP_DATE]
                //System.Drawing.Point vCellPoint06 = new System.Drawing.Point(9, 2);   //적요[REMARK]

                mPrinting.XLActiveSheet("SourceTab1"); //셀에 문자를 넣기 위해 쉬트 선택

                ////전표번호[GL_NUM]
                //vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("GL_NUM"));
                //if (vObject != null)
                //{
                //    vObject = string.Format("전표번호 : {0}", vObject);
                //    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vObject);
                //}
                //else
                //{
                //    vObject = null;
                //    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vObject);
                //}

                ////작성일자[SLIP_DATE]
                //vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("BATCH_DATE"));
                //if (vObject != null)
                //{
                //    vObject = ConvertDate(vObject);
                //    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, vObject);
                //}
                //else
                //{
                //    vObject = null;
                //    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, vObject);
                //}

                ////작성부서명[DEPT_NAME]
                //vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("DEPT_NAME"));
                //if (vObject != null)
                //{
                //    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, vObject);
                //}
                //else
                //{
                //    vObject = null;
                //    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, vObject);
                //}

                ////작성자 이름[PERSON_NAME]
                //vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("PERSON_NAME"));
                //if (vObject != null)
                //{
                //    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, vObject);
                //}
                //else
                //{
                //    vObject = null;
                //    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, vObject);
                //}

                //전표일자[GL_DATE]
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("GL_DATE"));
                if (vObject != null)
                {
                    vObject = ConvertDate(vObject);
                    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, vObject);
                }

                ////적요[REMARK]
                //string vText = string.Empty;
                //vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("REMARK"));
                //if (vObject != null)
                //{
                //    bool isConvert = vObject is string;
                //    if (isConvert == true)
                //    {
                //        vText = vObject as string;
                //        bool isNull = string.IsNullOrEmpty(vText.Trim());
                //        if (isNull != true)
                //        {
                //            vText = string.Format("내역 : {0}", vObject);
                //        }
                //    }
                //    vObject = vText;
                //    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, vObject);
                //}
                //else
                //{
                //    vObject = null;
                //    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, vObject);
                //}
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #endregion;

        #region ----- Line SLIP Methods ----

        #region ----- Array Set ----

        private void SetArray(out string[] pDBColumn, out int[] pXLColumn)
        {
            pDBColumn = new string[10];
            pXLColumn = new int[10];

            string vDBColumn01 = "SEQ_NO";
            string vDBColumn02 = "TAX_REG_NO";
            string vDBColumn03 = "VENDOR_NAME";
            string vDBColumn04 = "BANK_CODE";
            string vDBColumn05 = "BANK_NAME";
            string vDBColumn06 = "ACCOUNT_HOLDER";
            string vDBColumn07 = "BANK_ACCOUNT_NUM";
            string vDBColumn08 = "CURRENCY_CODE";
            string vDBColumn09 = "AMOUNT";
            string vDBColumn10 = "REMARK";

            pDBColumn[0] = vDBColumn01;  //SEQ_NO
            pDBColumn[1] = vDBColumn02;  //TAX_REG_NO
            pDBColumn[2] = vDBColumn03;  //VENDOR_NAME
            pDBColumn[3] = vDBColumn04;  //BANK_CODE
            pDBColumn[4] = vDBColumn05;  //BANK_NAME
            pDBColumn[5] = vDBColumn06;  //ACCOUNT_HOLDER
            pDBColumn[6] = vDBColumn07;  //BANK_ACCOUNT_NUM
            pDBColumn[7] = vDBColumn08;  //CURRENCY_CODE
            pDBColumn[8] = vDBColumn09;  //AMOUNT
            pDBColumn[9] = vDBColumn10;  //REMARK

            int vXLColumn01 = 1;         //SEQ_NO
            int vXLColumn02 = 3;         //TAX_REG_NO
            int vXLColumn03 = 11;        //VENDOR_NAME
            int vXLColumn04 = 23;        //BANK_CODE
            int vXLColumn05 = 27;        //BANK_NAME
            int vXLColumn06 = 33;        //ACCOUNT_HOLDER
            int vXLColumn07 = 39;         //BANK_ACCOUNT_NUM
            int vXLColumn08 = 48;        //CURRENCY_CODE
            int vXLColumn09 = 52;        //AMOUNT
            int vXLColumn10 = 59;        //REMARK

            pXLColumn[0] = vXLColumn01;  //SEQ_NO
            pXLColumn[1] = vXLColumn02;  //TAX_REG_NO
            pXLColumn[2] = vXLColumn03;  //VENDOR_NAME
            pXLColumn[3] = vXLColumn04;  //BANK_CODE
            pXLColumn[4] = vXLColumn05;  //BANK_NAME
            pXLColumn[5] = vXLColumn06;  //ACCOUNT_HOLDER
            pXLColumn[6] = vXLColumn07;  //BANK_ACCOUNT_NUM
            pXLColumn[7] = vXLColumn08;  //CURRENCY_CODE
            pXLColumn[8] = vXLColumn09;  //AMOUNT
            pXLColumn[9] = vXLColumn10;  //REMARK
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

        #endregion;

        #region ----- XlLine Methods -----

        private int XlLine(System.Data.DataRow pRow, int pPrintingLine, string[] pDBColumn, int[] pXLColumn)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            string vColumnName1= string.Empty;
            int vXLColumnIndex = 0;

            string vConvertString1 = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert1 = false;

            try
            {
                //[SEQ_NO]
                vColumnName1 = pDBColumn[0];
                vXLColumnIndex = pXLColumn[0];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                //[TAX_REG_NO]
                vColumnName1 = pDBColumn[1];
                vXLColumnIndex = pXLColumn[1];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);

                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                //[VENDOR_NAME]
                vColumnName1 = pDBColumn[2];
                vXLColumnIndex = pXLColumn[2];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);

                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                //[BANK_CODE]
                vColumnName1 = pDBColumn[3];
                vXLColumnIndex = pXLColumn[3];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                //[BANK_NAME]
                vColumnName1 = pDBColumn[4];
                vXLColumnIndex = pXLColumn[4];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                
                //[ACCOUNT_HOLDER]
                vColumnName1 = pDBColumn[5];
                vXLColumnIndex = pXLColumn[5];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                //[BANK_ACCOUNT_NUM]
                vColumnName1 = pDBColumn[6];
                vXLColumnIndex = pXLColumn[6];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                //[CURRENCY_CODE]
                vColumnName1 = pDBColumn[7];
                vXLColumnIndex = pXLColumn[7];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                //[AMOUNT]
                vColumnName1 = pDBColumn[8];
                vXLColumnIndex = pXLColumn[8];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[REMARK]
                vColumnName1 = pDBColumn[9];
                vXLColumnIndex = pXLColumn[9];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                vXLine = vXLine+1;

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }


            pPrintingLine = vXLine;

            return pPrintingLine;
        }

        #endregion;

        #region ----- Sum Write Methods -----

        private void SumWrite(int pPrintingLine)
        {
            //if (mFirstPagePrinted == false)
            //{
            //    mPrinting.XLActiveSheet("SourceTab1");
            //    if (mPrintingLineSTART1 != pPrintingLine) //66라인의 1페이지 출력물에서 2페이지 준비 때문에 미리 표시된 쉬트에 대해 Skip 되도록 하기위해 비교
            //    {
            //        //[합계]
            //        mPrinting.XLSetCell(53, 2, "합계");
            //        string vDRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
            //        mPrinting.XLSetCell(53, 12, vDRAmount);
            //        string vCRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
            //        mPrinting.XLSetCell(53, 18, vCRAmount);

            //        XlLineClear(pPrintingLine);
            //    }
            //    else
            //    {
            //        //[합계]
            //        mPrinting.XLSetCell(53, 2, "합계");
            //        string vDRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
            //        mPrinting.XLSetCell(53, 12, vDRAmount);
            //        string vCRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
            //        mPrinting.XLSetCell(53, 18, vCRAmount);
            //    }
            //}
            //else
            //{
            //    mPrinting.XLActiveSheet("SourceTab2");
            //    if (mPrintingLineSTART1 != pPrintingLine) //66라인의 1페이지 출력물에서 2페이지 준비 때문에 미리 표시된 쉬트에 대해 Skip 되도록 하기위해 비교
            //    {
            //        //[합계]
            //        mPrinting.XLSetCell(59, 2, "합계");
            //        string vDRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
            //        mPrinting.XLSetCell(59, 12, vDRAmount);
            //        string vCRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
            //        mPrinting.XLSetCell(59, 18, vCRAmount);

            //        XlLineClear(pPrintingLine);
            //    }
            //    else
            //    {
            //        //[합계]
            //        mPrinting.XLSetCell(59, 2, "합계");
            //        string vDRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
            //        mPrinting.XLSetCell(59, 12, vDRAmount);
            //        string vCRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
            //        mPrinting.XLSetCell(59, 18, vCRAmount);
            //    }
            //}
        }

        #endregion;

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;
            mFirstPagePrinted = false;

            string[] vDBColumn;
            int[] vXLColumn;

            mDR_AMOUNT = 0;
            mCR_AMOUNT = 0;

            int vPrintingLine = mPrintingLineSTART1;

            try
            {
                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    ComputeLastPageNumber(vTotalRow);

                    int vCountRow = 0;

                    SetArray(out vDBColumn, out vXLColumn);

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vPrintingLine = XlLine(vRow, vPrintingLine, vDBColumn, vXLColumn);

                        if (vTotalRow == vCountRow)
                        {
                            SumWrite(vPrintingLine);

                            mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM);
                            //XlAllContentClear(mPrinting); //지움
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                if (mFirstPagePrinted == false)
                                {
                                    vPrintingLine = mPrintingLineSTART1;
                                }
                                else
                                {
                                    vPrintingLine = mPrintingLineSTART2;
                                }
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

            return mPageNumber;
        }

        #endregion;
        public int LineWrite2(InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;
            mFirstPagePrinted = false;

            string[] vDBColumn;
            int[] vXLColumn;

            mDR_AMOUNT = 0;
            mCR_AMOUNT = 0;

            int vPrintingLine = mPrintingLineSTART1;

            try
            {
                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    ComputeLastPageNumber(vTotalRow);

                    int vCountRow = 0;

                    SetArray(out vDBColumn, out vXLColumn);

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;
                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vPrintingLine = XlLine(vRow, vPrintingLine, vDBColumn, vXLColumn);

                        if (vTotalRow == vCountRow)
                        {
                            SumWrite(vPrintingLine);

                            //mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM);
                            //XlAllContentClear(mPrinting); //지움
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                if (mFirstPagePrinted == false)
                                {
                                    vPrintingLine = mPrintingLineSTART1;
                                }
                                else
                                {
                                    vPrintingLine = mPrintingLineSTART2;
                                }
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

            return mPageNumber;
        }

        #region ----- Last Page Number Compute Methods ----

        private void ComputeLastPageNumber(int pTotalRow)
        {
            int vRow = 0;
            mPageTotalNumber = 1;

            if (pTotalRow > 12)
            {
                vRow = pTotalRow - 12;
                mPageTotalNumber = vRow / 18;
                mPageTotalNumber = (vRow % 18) == 0 ? (mPageTotalNumber + 1) : (mPageTotalNumber + 2);
            }
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine)
        {
            if (mFirstPagePrinted == false)
            {
                if (mPrintingLineEND1 < pPrintingLine)
                {
                    mIsNewPage = true;
                    mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM);

                    XlAllContentClear(mPrinting);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
            else
            {
                if (mPrintingLineEND2 < pPrintingLine)
                {
                    mIsNewPage = true;
                    mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM);

                    XlAllContentClear(mPrinting);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            if (mFirstPagePrinted == false)
            {
                mPageNumber++; //페이지 번호
                string vPageNumberText = string.Format("Page {0} of {1}", mPageNumber, mPageTotalNumber);
                mPrinting.XLActiveSheet("SourceTab1"); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.
                mPrinting.XLSetCell(41, 62, vPageNumberText); //페이지 번호, XLcell[행, 열]

                int vCopyPrintingRowSTART = vCopySumPrintingLine;
                vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
                int vCopyPrintingRowEnd = vCopySumPrintingLine;
                pPrinting.XLActiveSheet("SourceTab1");
                object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
                pPrinting.XLActiveSheet("Destination");
                object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
                pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

                mFirstPagePrinted = true;
            }
            else
            {
                mPageNumber++; //페이지 번호
                string vPageNumberText = string.Format("Page {0} of {1}", mPageNumber, mPageTotalNumber);
                mPrinting.XLActiveSheet("SourceTab1"); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.
                mPrinting.XLSetCell(41, 62, vPageNumberText); //페이지 번호, XLcell[행, 열]

                int vCopyPrintingRowSTART = vCopySumPrintingLine;
                vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
                int vCopyPrintingRowEnd = vCopySumPrintingLine;
                pPrinting.XLActiveSheet("SourceTab1");
                object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
                pPrinting.XLActiveSheet("Destination");
                object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
                pPrinting.XLCopyRange(vRangeSource, vRangeDestination);
            }

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND);
        }

        public void PreView(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;
        public void PDF(string pSaveFileName)
        {
            try
            {
                System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

                int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
                vMaxNumber = vMaxNumber + 1;
                string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);

                vSaveFileName = string.Format("{0}\\{1}", vWallpaperFolder, vSaveFileName);

                //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
                //vMaxNumber = vMaxNumber + 1;

                //pSaveFileName = pSaveFileName + vMaxNumber;

               /* bool isSuccess =*/ mPrinting.XLSaveAs_PDF(vSaveFileName);  // DELETED, BY MJSHIN
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;
    }
}