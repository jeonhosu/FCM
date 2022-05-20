using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace FCMF0212
{
    public class XLPrinting
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        // 쉬트명 정의.
        private string mTargetSheet = "Destination";
        private string mSourceSheet1 = "SourceTab1";
        private string mSourceSheet2 = "SourceTab2";

        private int mPageTotalNumber = 0;
        private int mPageNumber = 0;
        
        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 0;
        private int mCopy_StartRow = 0;
        private int mCopy_EndCol = 0;
        private int mCopy_EndRow = 0;
        private int mPrintingLastRow = 0;  //최종 인쇄 라인.

        private int mCurrentRow = 0;       //현재 인쇄되는 row 위치.
        private int mDefaultEndPageRow = 0;    // 페이지 증가후 PageCount 기본값.
        private int mDefaultPageRow = 5;    // 페이지 증가후 PageCount 기본값.

        private int mCountLinePrinting = 0; //엑셀 라인 Seq

        private decimal mDR_AMOUNT = 0; //차변합계
        private decimal mCR_AMOUNT = 0; //대변합계
        private decimal mCURR_DR_AMOUNT = 0; //차변합계
        private decimal mCURR_CR_AMOUNT = 0; //대변합계 

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
            pPrinting.XLActiveSheet("SourceTab1");

            //int vStartRow = mPrintingLineSTART1;
            //int vStartCol = mCopyColumnSTART;
            //int vEndRow = mPrintingLineEND1 + 5;
            //int vEndCol = mCopyColumnEND - 1;

            //mPrinting.XLSetCell(vStartRow, vStartCol, vEndRow, vEndCol, vObject);
        
        }

        #endregion;

        #region ----- Line Clear All Methods ----

        private void XlLineClear(int pPrintingLine)
        {
            mPrinting.XLActiveSheet("SourceTab1");

            //int vStartRow = pPrintingLine + 1;
            //int vStartCol = mCopyColumnSTART + 1;
            //int vEndRow = mPrintingLineEND1 - 4;
            //int vEndCol = mCopyColumnEND - 1;

            //if (vStartRow > vEndRow)
            //{
            //    vStartRow = vEndRow; //시작하는 행이 계산후, 끝나는 행 보다 값이 커지므로, 끝나는 행 값을 줌
            //}

            //mPrinting.XL_LineClearInSide(vStartRow, vStartCol, vEndRow, vEndCol);
            //mPrinting.XL_LineClearInSide(vEndRow + 2, vStartCol, vEndRow, vEndCol);
        
        }

        #endregion;

        #region ----- Excel Wirte [Header] Methods ----

        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pLOCAL_DATE)
        {
            object vObject = null;
            string vString = string.Empty;
            int vLine = 4;

            // 쉬트명 정의.
            mTargetSheet = "Destination";
            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(2, 6, vString);

                //작성부서명[DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(2, 12, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pAdapter.CurrentRow["PERSON_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(2, 57, vString);

                //인쇄일시[PRINTED DATE]
                vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(34, 1, vString);


                ///////////////////////////////////////////////////////////////////////
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택

                //작성일자[SLIP_DATE]
                vObject = pAdapter.CurrentRow["GL_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 32, vString);

                vLine = vLine + 1;
                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["GL_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);                    
                }
                else
                {
                    vString = string.Empty;               
                }
                mPrinting.XLSetCell(vLine, 32, vString);

                vLine = 8;
                //작성부서명[DEPT_CODE DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);  
                }
                else
                {
                    vString = string.Empty;  
                }
                mPrinting.XLSetCell(vLine, 6, vString);

                //작성부서명[DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);  
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 12, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pAdapter.CurrentRow["PERSON_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);  
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 57, vString);

                //인쇄일시[PRINTED DATE]
                vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(34, 1, vString);
 
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        public void HeaderWrite_BSK(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pSOB_DESC, object pLOCAL_DATE)
        {
            object vObject = null;
            string vString = string.Empty;

            // 쉬트명 정의.
            mTargetSheet = "Destination";
            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                 
                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["SLIP_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("[{0}]", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(1, 37, vString);

                //인쇄일시[PRINTED DATE]
                vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(56, 36, vString);


                ///////////////////////////////////////////////////////////////////////
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택

                //회계단위
                mPrinting.XLSetCell(8, 6, pSOB_DESC);

                ////작성부서명[DEPT_CODE DEPT_NAME]
                //vObject = pAdapter.CurrentRow["DEPT_CODE"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(2, 6, vString);

                //작성부서명[DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(10, 6, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pAdapter.CurrentRow["PERSON_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(12, 6, vString);

                //전표유형
                vObject = pAdapter.CurrentRow["SLIP_TYPE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(8, 20, vString);


                //작성일자[SLIP_DATE]
                vObject = pAdapter.CurrentRow["SLIP_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", iDate.ISGetDate(vObject));
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(10, 20, vString);

                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["SLIP_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(10, 25, vString);

                //전표일자[GL_DATE]
                vObject = pAdapter.CurrentRow["GL_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(12, 20, vString);

                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["GL_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(12, 25, vString);

                //인쇄일시[PRINTED DATE]
                vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                if (iString.ISNull(vObject) != string.Empty)
                { 
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(56, 36, vString);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        public void HeaderWrite_SEK(InfoSummit.Win.ControlAdv.ISDataAdapter pHeader, object pLOCAL_DATE)
        {// 헤더 인쇄.
            object vObject;
            int vXLine = 0;
            int vXLColumn = 0;

            // 쉬트명 정의.
            mTargetSheet = "Sheet1";
            mSourceSheet1 = "Source1";
            mSourceSheet2 = "Source2";

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                //전표 유형
                vXLine = 1;
                vXLColumn = 2;
                vObject = pHeader.CurrentRow["SLIP_TYPE_NAME"];
                if (vObject != DBNull.Value)
                {
                    vObject = string.Format("{0}", vObject);
                }
                else
                {
                    vObject = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //전표번호[SLIP_NUM]
                vXLine = 7;
                vXLColumn = 2;
                vObject = pHeader.CurrentRow["SLIP_NUM"];
                if (vObject != DBNull.Value)
                {
                    vObject = string.Format("발의번호 : {0}", vObject);
                }
                else
                {
                    vObject = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //발의일자[SLIP_DATE]
                vXLine = 9;
                vXLColumn = 7;
                vObject = pHeader.CurrentRow["SLIP_DATE"];
                if (vObject != DBNull.Value)
                {
                    if (iDate.ISDate(vObject) == true)
                    {
                        if (iDate.ISGetDate(vObject).ToShortDateString() == "0001-01-01")
                        {
                            vObject = iString.ISNull(vObject);
                        }
                        else
                        {
                            vObject = iDate.ISGetDate(vObject).ToShortDateString();
                        }
                    }
                    else
                    {
                        vObject = iString.ISNull(vObject);
                    }
                }
                else
                {
                    vObject = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //발의부서[DEPT_NAME]
                vXLine = 11;
                vXLColumn = 7;
                vObject = pHeader.CurrentRow["DEPT_NAME"];
                if (vObject != DBNull.Value)
                {

                }
                else
                {
                    vObject = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //발의자[PERSON_NAME]
                vXLine = 13;
                vXLColumn = 7;
                vObject = pHeader.CurrentRow["PERSON_NAME"];
                if (vObject != DBNull.Value)
                {
                }
                else
                {
                    vObject = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //수신[재경팀]
                vXLine = 15;
                vXLColumn = 7;
                vObject = "자금팀";
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                ////금액[SUB_REMARK]
                //vXLine = 17;
                //vXLColumn = 7;
                //vObject = pHeader.CurrentRow["SUB_REMARK"];
                //if (vObject != DBNull.Value)
                //{
                //}
                //else
                //{
                //    vObject = null;
                //}
                //mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //제목[REMARK]
                vXLine = 19;
                vXLColumn = 2;
                vObject = pHeader.CurrentRow["REMARK"];
                if (vObject != System.DBNull.Value)
                {
                    vObject = string.Format("제  목 : {0}", vObject);
                }
                else
                {
                    vObject = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                ////내역[SUBSTANCE]
                //vXLine = 23;
                //vXLColumn = 2;
                //string vContent = string.Empty;
                //vObject = pHeader.CurrentRow["SUBSTANCE"];
                //if (vObject != System.DBNull.Value)
                //{
                //    bool isConvert = vObject is string;
                //    if (isConvert == true)
                //    {
                //        vContent = vObject as string;
                //        bool isNull = string.IsNullOrEmpty(vContent.Trim());
                //        if (isNull != true)
                //        {
                //            //byte b_CR_Character = 0x0d; //CR
                //            //byte b_SP_Character = 0x20; //Space
                //            //char vCharOld = (char)b_CR_Character;
                //            //char vCharNew = (char)b_SP_Character;
                //            //vContent = vContent.Replace(vCharOld, vCharNew);
                //            vContent = vContent.Replace("\r", "");
                //        }
                //    }
                //    vObject = vContent.ToString(); ;
                //}
                //else
                //{
                //    vObject = null;
                //}
                //mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //인쇄일시[PRINTED DATE]
                vXLine = 72;
                vXLColumn = 2;
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vObject = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vObject = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                ////////////////////////////////////////
                mPrinting.XLActiveSheet(mSourceSheet2);

                //인쇄일시[PRINTED DATE]
                vXLine = 72;
                vXLColumn = 2;
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vObject = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vObject = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        public void HeaderWrite_BHC(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pSOB_DESC, object pLOCAL_DATE)
        {
            object vObject = null;
            string vString = string.Empty;

            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";
            mTargetSheet = "Destination";

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택

                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["GL_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("[{0}]", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(1, 37, vString);

                //인쇄일시[PRINTED DATE]
                vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(56, 36, vString);


                ///////////////////////////////////////////////////////////////////////
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택

                //회계단위
                mPrinting.XLSetCell(8, 6, pSOB_DESC);

                ////작성부서명[DEPT_CODE DEPT_NAME]
                //vObject = pAdapter.CurrentRow["DEPT_CODE"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(2, 6, vString);

                //작성부서명[DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(10, 6, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pAdapter.CurrentRow["PERSON_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(12, 6, vString);

                //전표유형
                vObject = pAdapter.CurrentRow["SLIP_TYPE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(8, 20, vString);


                //작성일자[SLIP_DATE]
                vObject = pAdapter.CurrentRow["SLIP_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", iDate.ISGetDate(vObject));
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(10, 20, vString);

                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["SLIP_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(10, 25, vString);

                //전표일자[GL_DATE]
                vObject = pAdapter.CurrentRow["GL_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", iDate.ISGetDate(vObject));
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(12, 20, vString);

                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["GL_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(12, 25, vString);

                //인쇄일시[PRINTED DATE]
                vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(56, 36, vString);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        public void HeaderWrite_SIK(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pSOB_DESC
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pSlip_Approval_Line
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pSlip_Approval_Line_User)
        {
            object vObject = null;
            string vString = string.Empty;

            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";
            mTargetSheet = "Destination";

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택

                //회계단위
                mPrinting.XLSetCell(1, 2, pSOB_DESC);


                //회계팀 승인단계// 
                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_TITLE");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 32, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_1");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 34, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_2");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 38, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_3");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 42, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_4");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 46, vString);

                //현업부서 승인단계// 
                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_TITLE");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 2, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_1");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 4, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_2");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 8, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_3");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 12, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_4");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 16, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_5");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 20, vString);
                ///// 승인단계 타이틀 인쇄 종료 /////
                
                //전표유형
                vObject = pAdapter.CurrentRow["SLIP_TYPE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 8, vString);

                //작성일자[SLIP_DATE]
                vObject = pAdapter.CurrentRow["SLIP_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(7, 8, vString);

                //전표일자[GL_DATE]
                vObject = pAdapter.CurrentRow["GL_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(8, 8, vString);

                //작성번호[GL_NUM]
                vObject = pAdapter.CurrentRow["SLIP_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 21, vString);


                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["GL_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(7, 21, vString);

                //작성부서명[DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(8, 21, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pAdapter.CurrentRow["PERSON_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(9, 21, vString);

                //인쇄자
                vObject = pAdapter.CurrentRow["PRINT_USER_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(69, 2, vString);

                //인쇄일시[PRINTED DATE]
                vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(69, 36, vString);
                 
                ///////////////////////////////////////////////////////////////////////
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택

                //회계단위
                mPrinting.XLSetCell(1, 2, pSOB_DESC);


                //회계팀 승인단계// 
                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_TITLE");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 32, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_1");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 34, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_2");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 38, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_3");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 42, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_4");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 46, vString);

                //현업부서 승인단계// 
                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_TITLE");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 2, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_1");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 4, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_2");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 8, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_3");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 12, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_4");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 16, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_5");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 20, vString);
                ///// 승인단계 타이틀 인쇄 종료 /////
                
                //전표유형
                vObject = pAdapter.CurrentRow["SLIP_TYPE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 8, vString);

                //작성일자[SLIP_DATE]
                vObject = pAdapter.CurrentRow["SLIP_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(7, 8, vString);

                //전표일자[GL_DATE]
                vObject = pAdapter.CurrentRow["GL_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(8, 8, vString);

                //작성번호[GL_NUM]
                vObject = pAdapter.CurrentRow["SLIP_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 21, vString);


                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["GL_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(7, 21, vString);

                //작성부서명[DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(8, 21, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pAdapter.CurrentRow["PERSON_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(9, 21, vString);

                //인쇄자
                vObject = pAdapter.CurrentRow["PRINT_USER_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(69, 2, vString);

                //인쇄일시[PRINTED DATE]
                vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(69, 36, vString);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        public void HeaderWrite_SIK(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pSOB_DESC
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pPrint_Appr_Person_TOP
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pPrint_Appr_Person_BOTTOM
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pSlip_Approval_Line
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pSlip_Approval_Line_User)
        {
            object vObject = null;
            string vString = string.Empty;

            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";
            mTargetSheet = "Destination";

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택

                //회계단위
                mPrinting.XLSetCell(1, 2, pSOB_DESC);


                //회계팀 승인단계// 
                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_TITLE");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 32, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_1");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 34, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_2");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 38, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_3");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 42, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_4");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 46, vString);

                //현업부서 승인단계// 
                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_TITLE");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 2, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_1");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 4, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_2");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 8, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_3");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 12, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_4");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 16, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_5");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 20, vString);
                ///// 승인단계 타이틀 인쇄 종료 /////

                //전표유형
                vObject = pAdapter.CurrentRow["SLIP_TYPE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 8, vString);

                //작성일자[SLIP_DATE]
                vObject = pAdapter.CurrentRow["SLIP_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(7, 8, vString);

                //전표일자[GL_DATE]
                vObject = pAdapter.CurrentRow["GL_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(8, 8, vString);

                //작성번호[GL_NUM]
                vObject = pAdapter.CurrentRow["SLIP_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 21, vString);


                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["GL_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(7, 21, vString);

                //작성부서명[DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(8, 21, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pAdapter.CurrentRow["PERSON_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(9, 21, vString);

                //인쇄자
                vObject = pAdapter.CurrentRow["PRINT_USER_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(69, 2, vString);

                //인쇄일시[PRINTED DATE]
                vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(69, 36, vString);

                //승인 여부에 처리.
                //작성자
                if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_N")).Equals("Y"))
                {
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_N")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_N");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(65, 4, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(67, 4, 68, 7);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }

                    //검토1
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(65, 8, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(67, 8, 68, 11);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }

                    //검토2
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A1")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A1");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(65, 12, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(67, 12, 68, 15);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }

                    //확인
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A2")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A2");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(65, 16, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(67, 16, 68, 19);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }

                    //승인
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_B")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_B");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(65, 20, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(67, 20, 68, 23);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }

                    //회계승인
                    if (iString.ISNull(pPrint_Appr_Person_TOP.GetCommandParamValue("O_APPR_N")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_TOP.GetCommandParamValue("O_PRINT_NAME_N");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(7, 34, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(9, 34, 10, 37);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }
                }

                ///////////////////////////////////////////////////////////////////////
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택

                //회계단위
                mPrinting.XLSetCell(1, 2, pSOB_DESC);


                //회계팀 승인단계// 
                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_TITLE");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 32, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_1");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 34, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_2");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 38, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_3");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 42, vString);

                vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_4");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 46, vString);

                //현업부서 승인단계// 
                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_TITLE");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 2, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_1");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 4, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_2");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 8, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_3");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 12, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_4");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 16, vString);

                vObject = pSlip_Approval_Line_User.GetCommandParamValue("O_PRINT_5");
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(64, 20, vString);
                ///// 승인단계 타이틀 인쇄 종료 /////

                //전표유형
                vObject = pAdapter.CurrentRow["SLIP_TYPE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 8, vString);

                //작성일자[SLIP_DATE]
                vObject = pAdapter.CurrentRow["SLIP_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(7, 8, vString);

                //전표일자[GL_DATE]
                vObject = pAdapter.CurrentRow["GL_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(8, 8, vString);

                //작성번호[GL_NUM]
                vObject = pAdapter.CurrentRow["SLIP_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(6, 21, vString);


                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["GL_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(7, 21, vString);

                //작성부서명[DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(8, 21, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pAdapter.CurrentRow["PERSON_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(9, 21, vString);

                //인쇄자
                vObject = pAdapter.CurrentRow["PRINT_USER_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(69, 2, vString);

                //인쇄일시[PRINTED DATE]
                vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(69, 36, vString);

                //승인 여부에 처리.
                //작성자
                if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_N")).Equals("Y"))
                {
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_N")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_N");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(65, 4, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(67, 4, 68, 7);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }

                    //검토1
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(65, 8, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(67, 8, 68, 11);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }

                    //검토2
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A1")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A1");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(65, 12, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(67, 12, 68, 15);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }

                    //확인
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A2")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A2");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(65, 16, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(67, 16, 68, 19);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }

                    //승인
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_B")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_B");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(65, 20, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(67, 20, 68, 23);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }

                    //회계승인
                    if (iString.ISNull(pPrint_Appr_Person_TOP.GetCommandParamValue("O_APPR_N")).Equals("Y"))
                    {
                        vObject = pPrint_Appr_Person_TOP.GetCommandParamValue("O_PRINT_NAME_N");
                        if (iString.ISNull(vObject) != string.Empty)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(7, 34, vString);

                        //승인 이미지
                        mPrinting.XLActiveSheet("Sheet1");
                        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                        mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                        object vRangeDestination = mPrinting.XLGetRange(9, 34, 10, 37);
                        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }


        public void HeaderWrite_DKT(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pSOB_DESC
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pSlip_Approval_Line_TOP
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pSlip_Approval_Line_BOTTOM)
        {
            int vLine = 1;
            object vObject = null;
            string vString = string.Empty;

            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";
            mTargetSheet = "Destination";

            try
            {
                for (int r = 0; r < 2; r++)
                {
                    vLine = 1;

                    if (r.Equals(0))
                        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                    else
                        mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택

                    //회계단위
                    mPrinting.XLSetCell(vLine, 2, pSOB_DESC);


                    //회계팀 승인단계// 
                    vLine = 9;
                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_TITLE");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 20, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_1");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 22, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_2");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 26, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_3");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 30, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_4");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 34, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_5");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 38, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_6");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 42, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_7");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 46, vString);

                    //현업부서 승인단계// 
                    vLine = 64;
                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_TITLE");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 2, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_1");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 4, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_2");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 8, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_3");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 12, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_4");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 16, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_5");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 20, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_6");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 24, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_7");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 28, vString);
                    ///// 승인단계 타이틀 인쇄 종료 /////

                    //전표유형
                    vObject = pAdapter.CurrentRow["SLIP_TYPE_NAME"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(6, 8, vString);

                    //작성일자[SLIP_DATE]
                    vObject = pAdapter.CurrentRow["SLIP_DATE"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(7, 8, vString);

                    //전표일자[GL_DATE]
                    vObject = pAdapter.CurrentRow["GL_DATE"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(8, 8, vString);

                    //작성번호[GL_NUM]
                    vObject = pAdapter.CurrentRow["SLIP_NUM"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(10, 8, vString);

                    //전표번호[GL_NUM]
                    vObject = pAdapter.CurrentRow["GL_NUM"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(11, 8, vString);

                    //작성부서명[DEPT_NAME]
                    vObject = pAdapter.CurrentRow["DEPT_NAME"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(12, 8, vString);

                    //작성자 이름[PERSON_NAME]
                    vObject = pAdapter.CurrentRow["PERSON_NAME"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(13, 8, vString);

                    //인쇄자
                    vObject = pAdapter.CurrentRow["PRINT_USER_NAME"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(69, 2, vString);

                    //인쇄일시[PRINTED DATE]
                    vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(69, 36, vString);
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        public void HeaderWrite_DKT(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pSOB_DESC
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pPrint_Appr_Person_TOP
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pPrint_Appr_Person_BOTTOM
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pSlip_Approval_Line_TOP
                                    , InfoSummit.Win.ControlAdv.ISDataCommand pSlip_Approval_Line_BOTTOM)
        {
            int vLine = 1;
            object vObject = null;
            string vString = string.Empty;

            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";
            mTargetSheet = "Destination";

            try
            {
                for (int r = 0; r < 2; r++)
                {
                    vLine = 1;

                    if (r.Equals(0))
                        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                    else
                        mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택

                    //회계단위
                    mPrinting.XLSetCell(vLine, 2, pSOB_DESC);


                    //회계팀 승인단계// 
                    vLine = 9;
                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_TITLE");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 20, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_1");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 22, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_2");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 26, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_3");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 30, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_4");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 34, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_5");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 38, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_6");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 42, vString);

                    vObject = pSlip_Approval_Line_TOP.GetCommandParamValue("O_PRINT_7");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 46, vString);

                    //현업부서 승인단계// 
                    vLine = 64;
                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_TITLE");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 2, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_1");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 4, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_2");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 8, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_3");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 12, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_4");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 16, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_5");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 20, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_6");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 24, vString);

                    vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_7");
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(vLine, 28, vString);
                    ///// 승인단계 타이틀 인쇄 종료 /////

                    //전표유형
                    vObject = pAdapter.CurrentRow["SLIP_TYPE_NAME"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(6, 8, vString);

                    //작성일자[SLIP_DATE]
                    vObject = pAdapter.CurrentRow["SLIP_DATE"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(7, 8, vString);

                    //전표일자[GL_DATE]
                    vObject = pAdapter.CurrentRow["GL_DATE"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(8, 8, vString);

                    //작성번호[GL_NUM]
                    vObject = pAdapter.CurrentRow["SLIP_NUM"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(10, 8, vString);

                    //전표번호[GL_NUM]
                    vObject = pAdapter.CurrentRow["GL_NUM"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(11, 8, vString);

                    //작성부서명[DEPT_NAME]
                    vObject = pAdapter.CurrentRow["DEPT_NAME"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(12, 8, vString);

                    //작성자 이름[PERSON_NAME]
                    vObject = pAdapter.CurrentRow["PERSON_NAME"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(13, 8, vString);

                    //인쇄자
                    vObject = pAdapter.CurrentRow["PRINT_USER_NAME"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(69, 2, vString);

                    //인쇄일시[PRINTED DATE]
                    vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vString = string.Empty;
                    }
                    mPrinting.XLSetCell(69, 36, vString);

                    //상단 승인단계 여부에 처리.
                    //작성자
                    vLine = 10;
                    if (iString.ISNull(pPrint_Appr_Person_TOP.GetCommandParamValue("O_APPR_N")).Equals("Y"))
                    {
                        if (iString.ISNull(pPrint_Appr_Person_TOP.GetCommandParamValue("O_APPR_N")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_TOP.GetCommandParamValue("O_PRINT_NAME_N");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 22, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 22, vLine + 3, 25);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //검토1
                        if (iString.ISNull(pPrint_Appr_Person_TOP.GetCommandParamValue("O_APPR_A")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_TOP.GetCommandParamValue("O_PRINT_NAME_A");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 26, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 26, vLine + 3, 29);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //검토2
                        if (iString.ISNull(pPrint_Appr_Person_TOP.GetCommandParamValue("O_APPR_A1")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_TOP.GetCommandParamValue("O_PRINT_NAME_A1");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 30, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 30, vLine + 3, 33);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //검토3
                        if (iString.ISNull(pPrint_Appr_Person_TOP.GetCommandParamValue("O_APPR_A2")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_TOP.GetCommandParamValue("O_PRINT_NAME_A2");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 34, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 34, vLine + 3, 37);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //검토4
                        if (iString.ISNull(pPrint_Appr_Person_TOP.GetCommandParamValue("O_APPR_A3")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_TOP.GetCommandParamValue("O_PRINT_NAME_A3");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 38, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 38, vLine + 3, 41);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //검토5
                        if (iString.ISNull(pPrint_Appr_Person_TOP.GetCommandParamValue("O_APPR_A4")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_TOP.GetCommandParamValue("O_PRINT_NAME_A4");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 42, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 42, vLine + 3, 45);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //확정
                        if (iString.ISNull(pPrint_Appr_Person_TOP.GetCommandParamValue("O_APPR_B")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_TOP.GetCommandParamValue("O_PRINT_NAME_B");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 46, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 46, vLine + 3, 49);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }
                    } 

                    //하단 승인 단계 여부에 처리.
                    //작성자
                    vLine = 65;
                    if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_N")).Equals("Y"))
                    {
                        if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_N")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_N");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 4, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 4, vLine + 3, 7);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //검토1
                        if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 8, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 8, vLine + 3, 11);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //검토2
                        if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A1")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A1");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 12, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 12, vLine + 3, 15);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //검토3
                        if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A2")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A2");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 16, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 16, vLine + 3, 19);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //검토4
                        if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A3")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A3");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 20, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 20, vLine + 3, 23);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //검토5
                        if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_A4")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_A4");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 24, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 24, vLine + 3, 27);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }

                        //확정
                        if (iString.ISNull(pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_APPR_B")).Equals("Y"))
                        {
                            vObject = pPrint_Appr_Person_BOTTOM.GetCommandParamValue("O_PRINT_NAME_B");
                            if (iString.ISNull(vObject) != string.Empty)
                            {
                                vString = string.Format("{0}", vObject);
                            }
                            else
                            {
                                vString = string.Empty;
                            }
                            mPrinting.XLSetCell(vLine, 28, vString);

                            //승인 이미지
                            mPrinting.XLActiveSheet("Sheet1");
                            object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                            if (r.Equals(0))
                                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                            else
                                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                            object vRangeDestination = mPrinting.XLGetRange(vLine + 2, 28, vLine + 3, 31);
                            mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                        }
                    }
                }

                /////////////////////////////////////////////////////////////////////////
                //mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택


                ////회계단위
                //mPrinting.XLSetCell(vLine, 2, pSOB_DESC);


                ////회계팀 승인단계// 
                //vLine = 9;
                //vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_TITLE");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 20, vString);

                //vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_1");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 22, vString);

                //vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_2");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 26, vString);

                //vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_3");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 30, vString);

                //vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_4");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 34, vString);

                //vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_5");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 38, vString);

                //vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_6");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 42, vString);

                //vObject = pSlip_Approval_Line.GetCommandParamValue("O_PRINT_7");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 46, vString);

                ////현업부서 승인단계// 
                //vLine = 64;
                //vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_TITLE");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 2, vString);

                //vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_1");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 4, vString);

                //vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_2");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 8, vString);

                //vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_3");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 12, vString);

                //vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_4");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 16, vString);

                //vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_5");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 20, vString);

                //vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_6");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 24, vString);

                //vObject = pSlip_Approval_Line_BOTTOM.GetCommandParamValue("O_PRINT_7");
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vLine, 28, vString);
                /////// 승인단계 타이틀 인쇄 종료 /////

                ////전표유형
                //vObject = pAdapter.CurrentRow["SLIP_TYPE_NAME"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(6, 8, vString);

                ////작성일자[SLIP_DATE]
                //vObject = pAdapter.CurrentRow["SLIP_DATE"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(7, 8, vString);

                ////전표일자[GL_DATE]
                //vObject = pAdapter.CurrentRow["GL_DATE"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(8, 8, vString);

                ////작성번호[GL_NUM]
                //vObject = pAdapter.CurrentRow["SLIP_NUM"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(10, 8, vString);

                ////전표번호[GL_NUM]
                //vObject = pAdapter.CurrentRow["GL_NUM"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(11, 8, vString);

                ////작성부서명[DEPT_NAME]
                //vObject = pAdapter.CurrentRow["DEPT_NAME"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(12, 8, vString);

                ////작성자 이름[PERSON_NAME]
                //vObject = pAdapter.CurrentRow["PERSON_NAME"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(13, 8, vString);

                ////인쇄자
                //vObject = pAdapter.CurrentRow["PRINT_USER_NAME"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(69, 2, vString);

                ////인쇄일시[PRINTED DATE]
                //vObject = pAdapter.CurrentRow["PRINT_DATETIME"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(69, 36, vString);

                ////현업 승인 여부에 처리.
                ////작성자
                //if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_N"]).Equals("Y"))
                //{
                //    if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_N"]).Equals("Y"))
                //    {
                //        vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_N"];
                //        if (iString.ISNull(vObject) != string.Empty)
                //        {
                //            vString = string.Format("{0}", vObject);
                //        }
                //        else
                //        {
                //            vString = string.Empty;
                //        }
                //        mPrinting.XLSetCell(65, 4, vString);

                //        //승인 이미지
                //        mPrinting.XLActiveSheet("Sheet1");
                //        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                //        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                //        object vRangeDestination = mPrinting.XLGetRange(67, 4, 68, 7);
                //        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                //    }

                //    //검토1
                //    if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_A"]).Equals("Y"))
                //    {
                //        vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_A"];
                //        if (iString.ISNull(vObject) != string.Empty)
                //        {
                //            vString = string.Format("{0}", vObject);
                //        }
                //        else
                //        {
                //            vString = string.Empty;
                //        }
                //        mPrinting.XLSetCell(65, 8, vString);

                //        //승인 이미지
                //        mPrinting.XLActiveSheet("Sheet1");
                //        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                //        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                //        object vRangeDestination = mPrinting.XLGetRange(67, 8, 68, 11);
                //        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                //    }

                //    //검토2
                //    if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_A1"]).Equals("Y"))
                //    {
                //        vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_A1"];
                //        if (iString.ISNull(vObject) != string.Empty)
                //        {
                //            vString = string.Format("{0}", vObject);
                //        }
                //        else
                //        {
                //            vString = string.Empty;
                //        }
                //        mPrinting.XLSetCell(65, 12, vString);

                //        //승인 이미지
                //        mPrinting.XLActiveSheet("Sheet1");
                //        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                //        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                //        object vRangeDestination = mPrinting.XLGetRange(67, 12, 68, 15);
                //        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                //    }

                //    //검토3
                //    if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_A2"]).Equals("Y"))
                //    {
                //        vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_A2"];
                //        if (iString.ISNull(vObject) != string.Empty)
                //        {
                //            vString = string.Format("{0}", vObject);
                //        }
                //        else
                //        {
                //            vString = string.Empty;
                //        }
                //        mPrinting.XLSetCell(65, 16, vString);

                //        //승인 이미지
                //        mPrinting.XLActiveSheet("Sheet1");
                //        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                //        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                //        object vRangeDestination = mPrinting.XLGetRange(67, 16, 68, 19);
                //        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                //    }

                //    //검토4
                //    if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_A3"]).Equals("Y"))
                //    {
                //        vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_A3"];
                //        if (iString.ISNull(vObject) != string.Empty)
                //        {
                //            vString = string.Format("{0}", vObject);
                //        }
                //        else
                //        {
                //            vString = string.Empty;
                //        }
                //        mPrinting.XLSetCell(65, 20, vString);

                //        //승인 이미지
                //        mPrinting.XLActiveSheet("Sheet1");
                //        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                //        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                //        object vRangeDestination = mPrinting.XLGetRange(67, 20, 68, 23);
                //        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                //    }

                //    //확정
                //    if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_B"]).Equals("Y"))
                //    {
                //        vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_B"];
                //        if (iString.ISNull(vObject) != string.Empty)
                //        {
                //            vString = string.Format("{0}", vObject);
                //        }
                //        else
                //        {
                //            vString = string.Empty;
                //        }
                //        mPrinting.XLSetCell(65, 24, vString);

                //        //승인 이미지
                //        mPrinting.XLActiveSheet("Sheet1");
                //        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                //        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                //        object vRangeDestination = mPrinting.XLGetRange(67, 20, 68, 23);
                //        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                //    }

                //    //회계승인
                //    if (iString.ISNull(pPrint_Appr_Person.CurrentRow["APPR_C"]).Equals("Y"))
                //    {
                //        vObject = pPrint_Appr_Person.CurrentRow["PRINT_NAME_C"];
                //        if (iString.ISNull(vObject) != string.Empty)
                //        {
                //            vString = string.Format("{0}", vObject);
                //        }
                //        else
                //        {
                //            vString = string.Empty;
                //        }
                //        mPrinting.XLSetCell(7, 22, vString);

                //        //승인 이미지
                //        mPrinting.XLActiveSheet("Sheet1");
                //        object vAppr_Image_RangeSource = mPrinting.XLGetRange(1, 1, 2, 4);

                //        mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                //        object vRangeDestination = mPrinting.XLGetRange(9, 34, 10, 37);
                //        mPrinting.XLCopyRange(vAppr_Image_RangeSource, vRangeDestination);
                //    }
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
            pDBColumn = new string[8];
            pXLColumn = new int[8];

            string vDBColumn01 = "ACCOUNT_CODE";
            string vDBColumn02 = "ACCOUNT_DESC";
            string vDBColumn03 = "DR_AMOUNT";
            string vDBColumn04 = "CR_AMOUNT";
            string vDBColumn05 = "M_REFERENCE";
            string vDBColumn06 = "REMARK";
            string vDBColumn07 = "CUSTOMER_DESC";
            string vDBColumn08 = "DEPT_DESC";

            pDBColumn[0] = vDBColumn01;  //ACCOUNT_CODE
            pDBColumn[1] = vDBColumn02;  //ACCOUNT_DESC
            pDBColumn[2] = vDBColumn03;  //DR_AMOUNT
            pDBColumn[3] = vDBColumn04;  //CR_AMOUNT
            pDBColumn[4] = vDBColumn05;  //M_REFERENCE
            pDBColumn[5] = vDBColumn06;  //REMARK
            pDBColumn[6] = vDBColumn07;  //CUSTOMER_DESC
            pDBColumn[7] = vDBColumn08;  //DEPT_DESC

            int vXLColumn01 = 3;         //ACCOUNT_CODE
            int vXLColumn02 = 3;         //ACCOUNT_DESC
            int vXLColumn03 = 12;        //DR_AMOUNT
            int vXLColumn04 = 18;        //CR_AMOUNT
            int vXLColumn05 = 24;        //M_REFERENCE
            int vXLColumn06 = 24;        //REMARK
            int vXLColumn07 = 24;        //CUSTOMER_DESC
            int vXLColumn08 = 40;        //DEPT_DESC

            pXLColumn[0] = vXLColumn01;  //ACCOUNT_CODE
            pXLColumn[1] = vXLColumn02;  //ACCOUNT_DESC
            pXLColumn[2] = vXLColumn03;  //DR_AMOUNT
            pXLColumn[3] = vXLColumn04;  //CR_AMOUNT
            pXLColumn[4] = vXLColumn05;  //M_REFERENCE
            pXLColumn[5] = vXLColumn06;  //REMARK
            pXLColumn[6] = vXLColumn07;  //CUSTOMER_DESC
            pXLColumn[7] = vXLColumn08;  //DEPT_DESC
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

        private int XlLine(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString= string.Empty;     

            mCountLinePrinting++;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                mPrinting.XLSetCell(vXLine, 1, mCountLinePrinting);

                //[ACCOUNT_CODE]
                vObject = pRow["ACCOUNT_CODE"];  
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject); 
                }
                else
                {
                    vString = string.Empty; 
                }
                mPrinting.XLSetCell(vXLine, 6, vString);

                //[ACCOUNT_DESC]
                vObject = pRow["ACCOUNT_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 6, vString);

                //[M_REFERENCE]
                vObject = pRow["M_REFERENCE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 6, vString);

                //[DEPT CODE]
                vObject = pRow["DEPT_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vString);

                //[DEPT_NAME]
                vObject = pRow["DEPT_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vString);

                //[CURRENCY]
                vObject = pRow["CURRENCY_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 26, vString);

                //[EXCHANGE_RATE]
                vObject = pRow["EXCHANGE_RATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###.####}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 26, vString);

                //[CURR_DR_AMOUNT]
                vObject = pRow["CURR_DR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###.####}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 31, vString);
                mCURR_DR_AMOUNT = mCURR_DR_AMOUNT + iString.ISDecimaltoZero(vObject, 0);

                //[REMARK]
                vObject = pRow["REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 31, vString);

                //[DR_AMOUNT]
                vObject = pRow["DR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 40, vString);
                mDR_AMOUNT = mDR_AMOUNT + iString.ISDecimaltoZero(vObject, 0);

                //[CURR_CR_AMOUNT]
                vObject = pRow["CURR_CR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###.####}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 49, vString);
                mCURR_CR_AMOUNT = mCURR_CR_AMOUNT + iString.ISDecimaltoZero(vObject, 0);

                //[CR_AMOUNT]
                vObject = pRow["CR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 58, vString);
                mCR_AMOUNT = mCR_AMOUNT + iString.ISDecimaltoZero(vObject, 0);

                //-------------------------------------------------------------------
                vXLine= vXLine + 3;
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

        private int XlLine_BSK(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString = string.Empty;

            mCountLinePrinting++;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                mPrinting.XLSetCell(vXLine, 2, mCountLinePrinting);

                //[ACCOUNT_CODE]
                vObject = pRow["ACCOUNT_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[ACCOUNT_DESC]
                vObject = pRow["ACCOUNT_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 10, vString);

                //[DR_AMOUNT]
                vObject = pRow["DR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 26, vString);
                mDR_AMOUNT = mDR_AMOUNT + iString.ISDecimaltoZero(vObject, 0);

                //[CR_AMOUNT]
                vObject = pRow["CR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vString);
                mCR_AMOUNT = mCR_AMOUNT + iString.ISDecimaltoZero(vObject, 0);

                //////////////////////
                vXLine++;
                //////////////////////

                //[DEPT_NAME]
                vObject = pRow["DEPT_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[EXCHANGE_RATE]
                vObject = pRow["EXCHANGE_RATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###.####}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vString);

                //[CURRENCY]
                vObject = pRow["CURRENCY_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 23, vString);

                //[CURR_DR_AMOUNT]
                vObject = pRow["CURR_DR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty && iString.ISNull(vObject) != "0")
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###.####}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 26, vString);
                mCURR_DR_AMOUNT = mCURR_DR_AMOUNT + iString.ISDecimaltoZero(vObject, 0);

                //[CURR_CR_AMOUNT]
                vObject = pRow["CURR_CR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty && iString.ISNull(vObject) != "0")
                {
                    vString = string.Format("{0:##,###,###,###,###,###,###.####}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vString);
                mCURR_CR_AMOUNT = mCURR_CR_AMOUNT + iString.ISDecimaltoZero(vObject, 0);

                //////////////////////
                vXLine++;
                //////////////////////

                ////[M_REFERENCE]
                //vObject = pRow["M_REFERENCE"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 2, 6, vString);

                //[REFERENCE1]
                vObject = pRow["MANAGEMENT1_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[REFERENCE2]
                vObject = pRow["MANAGEMENT2_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 19, vString);

                //[REFERENCE2]
                vObject = pRow["REFER1_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 34, vString);

                //////////////////////
                vXLine++;
                //////////////////////

                //[REFERENCE4]
                vObject = pRow["REFER2_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[REFERENCE5]
                vObject = pRow["REFER3_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 19, vString);

                //[REFERENCE6]
                vObject = pRow["REFER4_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 34, vString);

                //////////////////////
                vXLine++;
                //////////////////////

                //[REFERENCE7]
                vObject = pRow["REFER5_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[REFERENCE8]
                vObject = pRow["REFER6_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 19, vString);

                //[REFERENCE9]
                vObject = pRow["REFER7_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 34, vString);

                //////////////////////
                vXLine++;
                //////////////////////

                //[REMARK]
                vObject = pRow["REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[REFERENCE9]
                vObject = pRow["REFER8_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 34, vString);

                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            //////////////////////
            vXLine++;
            //////////////////////

            pPrintingLine = vXLine;

            return pPrintingLine;
        }

        public int XlLine_SEK(System.Data.DataRow pRow, int pPrintingLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //[ACCOUNT_CODE]
                vXLColumn = 2;
                vObject = pRow["ACCOUNT_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[DR_AMOUNT]
                vXLColumn = 14;
                vObject = pRow["DR_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                    mDR_AMOUNT = mDR_AMOUNT + iString.ISDecimaltoZero(vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[CR_AMOUNT]
                vXLColumn = 20;
                vObject = pRow["CR_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                    mCR_AMOUNT = mCR_AMOUNT + iString.ISDecimaltoZero(vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[M_REFERENCE]
                vXLColumn = 26;
                vObject = pRow["M_REFERENCE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------

                //[ACCOUNT_DESC]
                vXLColumn = 2;
                vObject = pRow["ACCOUNT_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------

                //[REMARK]
                vXLColumn = 26;
                vObject = pRow["REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }
        
        private int XlLine_SIK(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString = string.Empty;

            mCountLinePrinting++;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                //[SLIP_LINE_SEQ]
                vObject = pRow["SLIP_LINE_SEQ"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 2, vString);

                //[DEPT_NAME]
                vObject = pRow["DEPT_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[EXCHANGE_RATE]
                vObject = pRow["EXCHANGE_RATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0:##,###.####}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vString);

                //[CURRENCY]
                vObject = pRow["CURRENCY_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 24, vString);

                //[CURR_DR_AMOUNT]
                vObject = pRow["CURR_DR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty && iString.ISNull(vObject) != "0")
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 30, vString);

                //[CURR_CR_AMOUNT]
                vObject = pRow["CURR_CR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty && iString.ISNull(vObject) != "0")
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 40, vString);

                //////////////////////
                vXLine++;
                //////////////////////

                //[ACCOUNT_CODE]
                vObject = pRow["ACCOUNT_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[ACCOUNT_DESC]
                vObject = pRow["ACCOUNT_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vString);

                //[DR_AMOUNT]
                vObject = pRow["DR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 30, vString);

                //[CR_AMOUNT]
                vObject = pRow["CR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 40, vString);

                //////////////////////
                vXLine++;
                //////////////////////

                ////[M_REFERENCE]
                //vObject = pRow["M_REFERENCE"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 2, 6, vString);

                //[REFERENCE1]
                vObject = pRow["MANAGEMENT1_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[REFERENCE2]
                vObject = pRow["MANAGEMENT2_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 20, vString);

                //[REFERENCE2]
                vObject = pRow["REFER1_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 35, vString);

                //////////////////////
                vXLine++;
                //////////////////////

                //[REFERENCE4]
                vObject = pRow["REFER2_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[REFERENCE5]
                vObject = pRow["REFER3_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 20, vString);

                //[REFERENCE6]
                vObject = pRow["REFER4_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 35, vString);

                //////////////////////
                vXLine++;
                //////////////////////

                //[REFERENCE7]
                vObject = pRow["REFER5_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[REFERENCE8]
                vObject = pRow["REFER6_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 20, vString);

                //[REFERENCE9]
                vObject = pRow["REFER7_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 35, vString);

                //////////////////////
                vXLine++;
                //////////////////////

                //[REMARK]
                vObject = pRow["REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[REFERENCE9]
                vObject = pRow["REFER8_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 35, vString);

                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            //////////////////////
            vXLine++;
            //////////////////////

            pPrintingLine = vXLine;

            return pPrintingLine;
        }
         
        #endregion;

        #region ----- Sum Write Methods -----

        private void SumWrite(System.Data.DataRow pRow, int pPrintingLine)
        {
            mPrinting.XLActiveSheet(mTargetSheet);

            //PageNumber 인쇄//
            int vPageNumber = 34;
            int vLINE = 0;
            for (int r = 1; r <= mPageNumber; r++)
            {
                vLINE = vPageNumber * r;
                mPrinting.XLSetCell(vLINE, 29, string.Format("Page {0} of {1}", r, mPageNumber));

                if (r == mPageNumber)
                {
                    //
                }
                else 
                {
                    vLINE = vLINE - 1;
                    mPrinting.XLSetCell(vLINE, 1, "");
                }
            }

            //[합계]
            vLINE = vLINE - 1;
            mPrinting.XLSetCell(vLINE, 1, "SUM");
            object vObject;
            string vString = string.Empty;

            //[CURR_DR_AMOUNT]
            vObject = pRow["SUM_CURR_DR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 31, vString);

            //[DR_AMOUNT]
            vObject = pRow["SUM_DR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 40, vString);

            //[CURR_CR_AMOUNT]
            vObject = pRow["SUM_CURR_CR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 49, vString);

            //[CURR_CR_AMOUNT]
            vObject = pRow["SUM_CR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 58, vString);
             
        }

        private void SumWrite_BSK(int pPrintingLine)
        {
            mPrinting.XLActiveSheet(mTargetSheet);

            //PageNumber 인쇄//
            int vPageNumber = 56;
            int vLINE = 0;
            for (int r = 1; r <= mPageNumber; r++)
            {
                vLINE = vPageNumber * r;
                mPrinting.XLSetCell(vLINE, 14, string.Format("Page {0} of {1}", r, mPageNumber));

                if (r == mPageNumber)
                {
                    //
                }
                else
                {
                    vLINE = vLINE - 1;
                    mPrinting.XLSetCell(vLINE, 1, "");
                }
            }

            //[합계]
            vLINE = vLINE - 3;
            //mPrinting.XLSetCell(vLINE, 1, "SUM");
            string vAmount = string.Empty;

            //vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mCURR_DR_AMOUNT);
            //mPrinting.XLSetCell(vLINE, 31, vAmount);

            vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
            mPrinting.XLSetCell(vLINE, 29, vAmount);
            mPrinting.XLCellColorBrush(vLINE, 29, vLINE, 39, System.Drawing.Color.Silver);

            //vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mCURR_CR_AMOUNT);
            //mPrinting.XLSetCell(vLINE, 49, vAmount);

            vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
            mPrinting.XLSetCell(vLINE, 39, vAmount);

            //XlLineClear(pPrintingLine);
        }

        private void SumWrite_SEK(int pPrintingLine)
        {
            mPrinting.XLActiveSheet(mTargetSheet);

            //PageNumber 인쇄//
            int vPageNumber = 74;
            int vLINE = 0;
            for (int r = 1; r <= mPageNumber; r++)
            {
                vLINE = vPageNumber * r;
                vLINE = vLINE - 2;
                mPrinting.XLSetCell(vLINE, 38, string.Format("Page {0} of {1}", r, mPageNumber));

                if (r == mPageNumber)
                {
                    //
                }
                else
                {
                    //vLINE = vLINE - 1;
                    //mPrinting.XLSetCell(vLINE, 1, "");
                }
            }

            //[합계]
            vLINE = vLINE - 2;
            //mPrinting.XLSetCell(vLINE, 2, "합계");
            string vAmount = string.Empty;

            //vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mCURR_DR_AMOUNT);
            //mPrinting.XLSetCell(vLINE, 31, vAmount);

            vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
            mPrinting.XLSetCell(vLINE, 14, vAmount);
            //mPrinting.XLCellColorBrush(vLINE, 14, vLINE, 20, System.Drawing.Color.Silver);

            //vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mCURR_CR_AMOUNT);
            //mPrinting.XLSetCell(vLINE, 49, vAmount);

            vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
            mPrinting.XLSetCell(vLINE, 20, vAmount);

            //XlLineClear(pPrintingLine);
        }

        private void SumWrite_SIK(System.Data.DataRow pRow, int pPrintingLine)
        {
            mPrinting.XLActiveSheet(mTargetSheet);

            //PageNumber 인쇄//
            int vPageNumber = mCopy_EndRow;
            int vLINE = 0;
            int vAdd_Line = 11;
            for (int r = 1; r <= mPageNumber; r++)
            {
                mPrinting.XLSetCell(vLINE + vAdd_Line, 45, string.Format("{0} / {1}", r, mPageNumber));
                if (r < mPageNumber)
                {
                    mPrinting.XLSetCell(vLINE + 63, 21, "");
                    mPrinting.XLSetCell(vLINE + 64, 21, "");
                }
                vLINE = vPageNumber * r;
            }

            //[합계]
            //mPrinting.XLSetCell(vLINE, 1, "SUM");
            object vObject;
            string vString = string.Empty;

            vLINE = mCopy_EndRow * mPageNumber - 6;

            //[CURR_DR_AMOUNT]
            vObject = pRow["SUM_CURR_DR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 30, vString);

            //[CURR_CR_AMOUNT]
            vObject = pRow["SUM_CURR_CR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 40, vString);

            vLINE = vLINE + 1;

            //[DR_AMOUNT]
            vObject = pRow["SUM_DR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 30, vString);

            //[CURR_CR_AMOUNT]
            vObject = pRow["SUM_CR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 40, vString);
        }

        private void SumWrite_DKT(System.Data.DataRow pRow, int pPrintingLine)
        {
            mPrinting.XLActiveSheet(mTargetSheet);

            //PageNumber 인쇄//
            int vPageNumber = mCopy_EndRow;
            int vLINE = 0;
            int vAdd_Line = 15;
            for (int r = 1; r <= mPageNumber; r++)
            {
                mPrinting.XLSetCell(vLINE + vAdd_Line, 45, string.Format("{0} / {1}", r, mPageNumber));
                if (r < mPageNumber)
                {
                    mPrinting.XLSetCell(vLINE + 61, 25, "");
                    mPrinting.XLSetCell(vLINE + 62, 25, "");
                }
                vLINE = vPageNumber * r;
            }

            //[합계]
            //mPrinting.XLSetCell(vLINE, 1, "SUM");
            object vObject;
            string vString = string.Empty;

            vLINE = mCopy_EndRow * mPageNumber - 8;

            //[CURR_DR_AMOUNT]
            vObject = pRow["SUM_CURR_DR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 30, vString);

            //[CURR_CR_AMOUNT]
            vObject = pRow["SUM_CURR_CR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 40, vString);

            vLINE = vLINE + 1;

            //[DR_AMOUNT]
            vObject = pRow["SUM_DR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 30, vString);

            //[CURR_CR_AMOUNT]
            vObject = pRow["SUM_CR_AMOUNT"];
            if (iString.ISNull(vObject) != string.Empty)
            {
                vString = string.Format("{0}", vObject);
            }
            else
            {
                vString = string.Empty;
            }
            mPrinting.XLSetCell(vLINE, 40, vString);
        }

        #endregion;

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false; 
             
            // 쉬트명 정의.
            mTargetSheet = "Destination";
            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";

            mDR_AMOUNT = 0;
            mCR_AMOUNT = 0;
            mCURR_DR_AMOUNT = 0;
            mCURR_CR_AMOUNT = 0;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 66;
            mCopy_EndRow = 34;
            
            mPrintingLastRow = 30;  //최종 인쇄 라인.
            mCurrentRow = 12;
            int vPrintingLine = mCurrentRow;
            
            try
            {
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                { 
                    int vCountRow = 0;
                     
                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 3;

                        if (vTotalRow == vCountRow)
                        {
                            IsNewPage(vPrintingLine);
                            SumWrite(vRow, mCurrentRow);

                            //mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM);
                            //XlAllContentClear(mPrinting);
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow + 2;
                                vPrintingLine = mDefaultPageRow + 1;

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

        public int LineWrite_BSK(InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            // 쉬트명 정의.
            mTargetSheet = "Destination";
            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";

            mDR_AMOUNT = 0;
            mCR_AMOUNT = 0;
            mCURR_DR_AMOUNT = 0;
            mCURR_CR_AMOUNT = 0;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 49;
            mCopy_EndRow = 56;

            mDefaultEndPageRow = 4;
            mDefaultPageRow = 4;    // 페이지 증가후 PageCount 기본값.
            mPrintingLastRow = 47;  //최종 인쇄 라인.
            mCurrentRow = 17;
            int vPrintingLine = mCurrentRow;

            try
            {
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_BSK(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 6;

                        if (vTotalRow == vCountRow)
                        {
                            SumWrite_BSK(mCurrentRow);  //조건부 서식으로 합계 인쇄.

                            //mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM);
                            //XlAllContentClear(mPrinting);
                        }
                        else
                        {
                            IsNewPage_BSK(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow + mDefaultEndPageRow;
                                vPrintingLine = mDefaultPageRow + 1;

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

        public int LineWrite_SEK(InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            // 쉬트명 정의.
            mTargetSheet = "Sheet1";
            mSourceSheet1 = "Source1";
            mSourceSheet2 = "Source2";

            mDR_AMOUNT = 0;
            mCR_AMOUNT = 0;
            mCURR_DR_AMOUNT = 0;
            mCURR_CR_AMOUNT = 0;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 45;
            mCopy_EndRow = 74;

            mDefaultEndPageRow = 5;
            mDefaultPageRow = 3;    // 페이지 증가후 PageCount 기본값.
            mPrintingLastRow = 67;  //최종 인쇄 라인.
            mCurrentRow = 49;
            int vPrintingLine = mCurrentRow;

            try
            {
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_SEK(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 3;

                        if (vTotalRow == vCountRow)
                        {
                            SumWrite_SEK(mCurrentRow);  //조건부 서식으로 합계 인쇄.

                            //mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM);
                            //XlAllContentClear(mPrinting);
                        }
                        else
                        {
                            IsNewPage_BSK(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow + mDefaultEndPageRow;
                                vPrintingLine = mDefaultPageRow + 1;

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
        
        public int LineWrite_BHC(InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";
            mTargetSheet = "Destination";

            mDR_AMOUNT = 0;
            mCR_AMOUNT = 0;
            mCURR_DR_AMOUNT = 0;
            mCURR_CR_AMOUNT = 0;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 49;
            mCopy_EndRow = 56;

            mDefaultEndPageRow = 4;
            mDefaultPageRow = 4;    // 페이지 증가후 PageCount 기본값.
            mPrintingLastRow = 47;  //최종 인쇄 라인.
            mCurrentRow = 17;
            int vPrintingLine = mCurrentRow;

            try
            {
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_BSK(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 6;

                        if (vTotalRow == vCountRow)
                        {
                            SumWrite_BSK(mCurrentRow);  //조건부 서식으로 합계 인쇄.

                            //mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM);
                            //XlAllContentClear(mPrinting);
                        }
                        else
                        {
                            IsNewPage_BSK(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow + mDefaultEndPageRow;
                                vPrintingLine = mDefaultPageRow + 1;

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

        public int LineWrite_SIK(InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";
            mTargetSheet = "Destination";

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 69;

            mDefaultPageRow = 14;    // 페이지 증가후 PageCount 기본값.
            mPrintingLastRow = 57;  //최종 인쇄 라인.
            mDefaultEndPageRow = 7;  //종료 후 남머지 ROW수
            mCurrentRow = 15;       //현재 ROW.
            int vPrintingLine = mCurrentRow;

            try
            {
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_SIK(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 6;

                        if (vTotalRow == vCountRow)
                        {
                            SumWrite_SIK(vRow, mCurrentRow);  //조건부 서식으로 합계 인쇄.
                        }
                        else
                        {
                            IsNewPage_BSK(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow + mDefaultEndPageRow;
                                vPrintingLine = mDefaultPageRow + 1;
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


        public int LineWrite_DKT(InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab2";
            mTargetSheet = "Destination";

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 69;

            mDefaultPageRow = 18;    // 페이지 증가후 PageCount 기본값.
            mPrintingLastRow = 55;  //최종 인쇄 라인.
            mDefaultEndPageRow = 9;  //종료 후 남머지 ROW수
            mCurrentRow = 19;       //현재 ROW.
            int vPrintingLine = mCurrentRow;

            try
            {
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_SIK(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 6;

                        if (vTotalRow == vCountRow)
                        {
                            SumWrite_DKT(vRow, mCurrentRow);  //조건부 서식으로 합계 인쇄.
                        }
                        else
                        {
                            IsNewPage_BSK(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow + mDefaultEndPageRow;
                                vPrintingLine = mDefaultPageRow + 1;
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
            if (mPrintingLastRow < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2,  mCopyLineSUM);

                //XlAllContentClear(mPrinting);
            }
            else
            {
                mIsNewPage = false;
            }
            
        }

        private void IsNewPage_BSK(int pPrintingLine)
        {
            if (mPrintingLastRow < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM);

                //XlAllContentClear(mPrinting);
            }
            else
            {
                mIsNewPage = false;
            }

        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(XL.XLPrint pPrinting, string pSourceTab, int pCopySumPrintingLine)
        {
            mPageNumber++; //페이지 번호
            
            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(pSourceTab); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pSourceTab);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            int vCopyPrintingRowSTART = pCopySumPrintingLine; 
            
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, mCopy_StartCol, vCopyPrintingRowSTART + mCopy_EndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow; 
            mPrinting.XLHPageBreaks_Add(mPrinting.XLGetRange("A" + vCopySumPrintingLine));
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
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 2);
        }

        public void Save(string pSaveFileName)
        {
            if (pSaveFileName == string.Empty)
            {
                return;
            }

            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }

        //PDF Method//
        public void PDF(string pSaveFileName)
        {
            try
            {
                mPrinting.XLDeleteSheet(mSourceSheet1);
                mPrinting.XLDeleteSheet(mSourceSheet2); 
                mPrinting.XLDeleteSheet("Sheet1");
                bool isSuccess = mPrinting.XLSaveAs_PDF(pSaveFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #endregion;
         
    }
}