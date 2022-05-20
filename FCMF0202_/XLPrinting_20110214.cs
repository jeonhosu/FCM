using System;

namespace FCMF0202
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

        private int mPrintingLineSTART1 = 14; //라인 출력시 엑셀 시작 행 위치 지정
        private int mPrintingLineEND1 = 52;   //mPrintingLineSTART1 부터 실제 출력될 마지막 행 위치 지정

        private int mPrintingLineSTART2 = 5;  //라인 출력시 엑셀 시작 행 위치 지정
        private int mPrintingLineEND2 = 58;   //mPrintingLineSTART2 부터 실제 출력될 마지막 행 위치 지정

        private int mCopyLineSUM = 1;        //엑셀의 선택된 쉬트의 복사되어질 시작 행 위치
        private int mIncrementCopyMAX = 62;  //복사되어질 행의 범위

        private int mCopyColumnSTART = 1; //복사되어진 행 누적 수
        private int mCopyColumnEND = 46;  //엑셀의 선택된 쉬트의 복사되어질 끝 열 위치

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
            int vXLColumn01 = 2;
            int vXLColumn02 = 12;
            int vXLColumn03 = 18;
            int vXLColumn04 = 24;

            object vObject = null;
            int vPrintingLineEND = 0;

            if (mFirstPagePrinted == false)
            {
                pPrinting.XLActiveSheet("SourceTab1");

                //vPrintingLineEND = mPrintingLineEND1 + 2;
                //for (int vXLine = mPrintingLineSTART1; vXLine < vPrintingLineEND; vXLine++)
                //{
                //    pPrinting.XLSetCell(vXLine, vXLColumn01, vObject);
                //    pPrinting.XLSetCell(vXLine, vXLColumn02, vObject);
                //    pPrinting.XLSetCell(vXLine, vXLColumn03, vObject);
                //    pPrinting.XLSetCell(vXLine, vXLColumn04, vObject);
                //}

                int vStartRow = mPrintingLineSTART1;
                int vStartCol = mCopyColumnSTART + 1;
                int vEndRow = mPrintingLineEND1;
                int vEndCol = mCopyColumnEND - 1;

                mPrinting.XLSetCell(vStartRow, vStartCol, vEndRow, vEndCol, vObject);
            }
            else
            {
                pPrinting.XLActiveSheet("SourceTab2");

                //vPrintingLineEND = mPrintingLineEND2 + 2;
                //for (int vXLine = mPrintingLineSTART2; vXLine < vPrintingLineEND; vXLine++)
                //{
                //    pPrinting.XLSetCell(vXLine, vXLColumn01, vObject);
                //    pPrinting.XLSetCell(vXLine, vXLColumn02, vObject);
                //    pPrinting.XLSetCell(vXLine, vXLColumn03, vObject);
                //    pPrinting.XLSetCell(vXLine, vXLColumn04, vObject);
                //}

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
            int vPrintingLineEND = 0;
            if (mFirstPagePrinted == false)
            {
                mPrinting.XLActiveSheet("SourceTab1");

                //vPrintingLineEND = mPrintingLineEND1 - 1;
                //for (int vXLine = pPrintingLine; vXLine < vPrintingLineEND; vXLine++)
                //{
                //    mPrinting.XL_LineClear(vXLine, mCopyColumnSTART, mCopyColumnEND);
                //}

                int vStartRow = pPrintingLine + 1;
                int vStartCol = mCopyColumnSTART + 1;
                int vEndRow = mPrintingLineEND2 - 1;
                int vEndCol = mCopyColumnEND - 1;

                mPrinting.XL_LineClearInSide(vStartRow, vStartCol, vEndRow, vEndCol);
            }
            else
            {
                mPrinting.XLActiveSheet("SourceTab2");

                //vPrintingLineEND = mPrintingLineEND2 - 1;
                //for (int vXLine = pPrintingLine; vXLine < vPrintingLineEND; vXLine++)
                //{
                //    mPrinting.XL_LineClear(vXLine, mCopyColumnSTART, mCopyColumnEND);
                //}

                int vStartRow = pPrintingLine;
                int vStartCol = mCopyColumnSTART + 1;
                int vEndRow = mPrintingLineEND2;
                int vEndCol = mCopyColumnEND - 1;

                mPrinting.XL_LineClearInSide(vStartRow, vStartCol, vEndRow, vEndCol);
            }
        }

        #endregion;

        #region ----- Excel Wirte [Header] Methods ----

        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow)
        {
            object vObject = null;

            try
            {
                int vIndexRowCurrent = pIndexRow;

                int vIndexColumn_01 = pGrid.GetColumnToIndex("GL_NUM");          //전표번호[GL_NUM]
                int vIndexColumn_02 = pGrid.GetColumnToIndex("SLIP_DATE");       //기표일자[SLIP_DATE]
                int vIndexColumn_03 = pGrid.GetColumnToIndex("DEPT_NAME");       //작성부서명[DEPT_NAME]
                int vIndexColumn_04 = pGrid.GetColumnToIndex("PERSON_NAME");     //작성자 이름[PERSON_NAME]
                int vIndexColumn_05 = pGrid.GetColumnToIndex("GL_DATE");         //작성일자[GL_DATE]
                int vIndexColumn_06 = pGrid.GetColumnToIndex("REMARK");          //적요[REMARK]

                System.Drawing.Point vGridPoint01 = new System.Drawing.Point(vIndexRowCurrent, vIndexColumn_01);  //전표번호[GL_NUM]
                System.Drawing.Point vGridPoint02 = new System.Drawing.Point(vIndexRowCurrent, vIndexColumn_02);  //작성일자[GL_DATE]
                System.Drawing.Point vGridPoint03 = new System.Drawing.Point(vIndexRowCurrent, vIndexColumn_03);  //작성부서명[DEPT_NAME]
                System.Drawing.Point vGridPoint04 = new System.Drawing.Point(vIndexRowCurrent, vIndexColumn_04);  //작성자 이름[PERSON_NAME]
                System.Drawing.Point vGridPoint05 = new System.Drawing.Point(vIndexRowCurrent, vIndexColumn_05);  //기표일자[SLIP_DATE]
                System.Drawing.Point vGridPoint06 = new System.Drawing.Point(vIndexRowCurrent, vIndexColumn_06);  //적요[REMARK]

                System.Drawing.Point vCellPoint01 = new System.Drawing.Point(7, 2);   //전표번호[GL_NUM]
                System.Drawing.Point vCellPoint02 = new System.Drawing.Point(1, 31);  //발의일자[SLIP_DATE]
                System.Drawing.Point vCellPoint03 = new System.Drawing.Point(3, 31);  //발의부서명[DEPT_NAME]
                System.Drawing.Point vCellPoint04 = new System.Drawing.Point(5, 31);  //발의자 이름[PERSON_NAME]
                System.Drawing.Point vCellPoint05 = new System.Drawing.Point(7, 31);  //전표일자[SLIP_DATE]
                System.Drawing.Point vCellPoint06 = new System.Drawing.Point(9, 2);   //적요[REMARK]

                mPrinting.XLActiveSheet("SourceTab1"); //셀에 문자를 넣기 위해 쉬트 선택

                //전표번호[GL_NUM]
                vObject = pGrid.GetCellValue(vGridPoint01.X, vGridPoint01.Y);
                if (vObject != null)
                {
                    vObject = string.Format("전표번호 : {0}", vObject);
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vObject);
                }

                //작성일자[GL_DATE]
                vObject = pGrid.GetCellValue(vGridPoint02.X, vGridPoint02.Y);
                if (vObject != null)
                {
                    vObject = ConvertDate(vObject);
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, vObject);
                }

                //작성부서명[DEPT_NAME]
                vObject = pGrid.GetCellValue(vGridPoint03.X, vGridPoint03.Y);
                if (vObject != null)
                {
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, vObject);
                }

                //작성자 이름[PERSON_NAME]
                vObject = pGrid.GetCellValue(vGridPoint04.X, vGridPoint04.Y);
                if (vObject != null)
                {
                    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, vObject);
                }

                //기표일자[SLIP_DATE]
                vObject = pGrid.GetCellValue(vGridPoint05.X, vGridPoint05.Y);
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

                //적요[REMARK]
                string vText = string.Empty;
                vObject = pGrid.GetCellValue(vGridPoint06.X, vGridPoint06.Y);
                if (vObject != null)
                {
                    bool isConvert = vObject is string;
                    if (isConvert == true)
                    {
                        vText = vObject as string;
                        bool isNull = string.IsNullOrEmpty(vText.Trim());
                        if (isNull != true)
                        {
                            vText = string.Format("내역 : {0}", vObject);
                        }
                    }
                    vObject = vText;
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, vObject);
                }
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
            pDBColumn = new string[6];
            pXLColumn = new int[6];

            string vDBColumn01 = "ACCOUNT_CODE";
            string vDBColumn02 = "ACCOUNT_DESC";
            string vDBColumn03 = "DR_AMOUNT";
            string vDBColumn04 = "CR_AMOUNT";
            string vDBColumn05 = "M_REFERENCE";
            string vDBColumn06 = "REMARK";

            pDBColumn[0] = vDBColumn01;  //ACCOUNT_CODE
            pDBColumn[1] = vDBColumn02;  //ACCOUNT_DESC
            pDBColumn[2] = vDBColumn03;  //DR_AMOUNT
            pDBColumn[3] = vDBColumn04;  //CR_AMOUNT
            pDBColumn[4] = vDBColumn05;  //M_REFERENCE
            pDBColumn[5] = vDBColumn06;  //REMARK

            int vXLColumn01 = 2;         //ACCOUNT_CODE
            int vXLColumn02 = 2;         //ACCOUNT_DESC
            int vXLColumn03 = 12;        //DR_AMOUNT
            int vXLColumn04 = 18;        //CR_AMOUNT
            int vXLColumn05 = 24;        //M_REFERENCE
            int vXLColumn06 = 24;        //REMARK

            pXLColumn[0] = vXLColumn01;  //ACCOUNT_CODE
            pXLColumn[1] = vXLColumn02;  //ACCOUNT_DESC
            pXLColumn[2] = vXLColumn03;  //DR_AMOUNT
            pXLColumn[3] = vXLColumn04;  //CR_AMOUNT
            pXLColumn[4] = vXLColumn05;  //M_REFERENCE
            pXLColumn[5] = vXLColumn06;  //REMARK
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

        private bool IsConvertNumber(string pStringNumber, out decimal pConvertDecimal)
        {
            bool vIsConvert = false;
            pConvertDecimal = 0m;

            try
            {
                if (pStringNumber != null)
                {
                    decimal vIsConvertNum = decimal.Parse(pStringNumber);
                    pConvertDecimal = vIsConvertNum;
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
                //[ACCOUNT_CODE]
                vColumnName1 = pDBColumn[0];
                vXLColumnIndex = pXLColumn[0];
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

                //[DR_AMOUNT]
                vColumnName1 = pDBColumn[2];
                vXLColumnIndex = pXLColumn[2];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);

                    vConvertString1 = vConvertString1.Replace(",", "");
                    IsConvertNumber(vConvertString1, out vConvertDecimal);
                    mDR_AMOUNT = mDR_AMOUNT + vConvertDecimal;
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[CR_AMOUNT]
                vColumnName1 = pDBColumn[3];
                vXLColumnIndex = pXLColumn[3];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);

                    vConvertString1 = vConvertString1.Replace(",", "");
                    IsConvertNumber(vConvertString1, out vConvertDecimal);
                    mCR_AMOUNT = mCR_AMOUNT + vConvertDecimal;
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[M_REFERENCE]
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

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------

                //[ACCOUNT_DESC]
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

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------

                //[REMARK]
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

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
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
            if (mFirstPagePrinted == false)
            {
                mPrinting.XLActiveSheet("SourceTab1");
                if (mPrintingLineSTART1 != pPrintingLine) //66라인의 1페이지 출력물에서 2페이지 준비 때문에 미리 표시된 쉬트에 대해 Skip 되도록 하기위해 비교
                {
                    //[합계]
                    mPrinting.XLSetCell(53, 2, "합계");
                    string vDRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
                    mPrinting.XLSetCell(53, 12, vDRAmount);
                    string vCRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
                    mPrinting.XLSetCell(53, 18, vCRAmount);

                    XlLineClear(pPrintingLine);
                }
                else
                {
                    //[합계]
                    mPrinting.XLSetCell(53, 2, "합계");
                    string vDRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
                    mPrinting.XLSetCell(53, 12, vDRAmount);
                    string vCRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
                    mPrinting.XLSetCell(53, 18, vCRAmount);
                }
            }
            else
            {
                mPrinting.XLActiveSheet("SourceTab2");
                if (mPrintingLineSTART1 != pPrintingLine) //66라인의 1페이지 출력물에서 2페이지 준비 때문에 미리 표시된 쉬트에 대해 Skip 되도록 하기위해 비교
                {
                    //[합계]
                    mPrinting.XLSetCell(59, 2, "합계");
                    string vDRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
                    mPrinting.XLSetCell(59, 12, vDRAmount);
                    string vCRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
                    mPrinting.XLSetCell(59, 18, vCRAmount);

                    XlLineClear(pPrintingLine);
                }
                else
                {
                    //[합계]
                    mPrinting.XLSetCell(59, 2, "합계");
                    string vDRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
                    mPrinting.XLSetCell(59, 12, vDRAmount);
                    string vCRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
                    mPrinting.XLSetCell(59, 18, vCRAmount);
                }
            }
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
                int vTotalRow = pData.OraSelectData.Rows.Count;
                if (vTotalRow > 0)
                {
                    ComputeLastPageNumber(vTotalRow);

                    int vCountRow = 0;

                    SetArray(out vDBColumn, out vXLColumn);

                    foreach (System.Data.DataRow vRow in pData.OraSelectData.Rows)
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
                            XlAllContentClear(mPrinting);
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

        #region ----- Last Page Number Compute Methods ----

        private void ComputeLastPageNumber(int pTotalRow)
        {
            int vRow = 0;
            mPageTotalNumber = 1;

            if (pTotalRow > 13)
            {
                vRow = pTotalRow - 13;
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
                mPrinting.XLSetCell(62, 14, vPageNumberText); //페이지 번호, XLcell[행, 열]

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
                mPrinting.XLActiveSheet("SourceTab2"); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.
                mPrinting.XLSetCell(62, 14, vPageNumberText); //페이지 번호, XLcell[행, 열]

                int vCopyPrintingRowSTART = vCopySumPrintingLine;
                vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
                int vCopyPrintingRowEnd = vCopySumPrintingLine;
                pPrinting.XLActiveSheet("SourceTab2");
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

        #endregion;

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