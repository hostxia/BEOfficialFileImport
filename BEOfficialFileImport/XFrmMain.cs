using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BEOfficialFileImport.DBUtility;
using DevExpress.Utils;
using DevExpress.XtraEditors;
using DevExpress.XtraPrinting.Native;
using DevExpress.XtraSplashScreen;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Image = System.Drawing.Image;

namespace BEOfficialFileImport
{
    public partial class XFrmMain : DevExpress.XtraEditors.XtraForm
    {
        public XFrmMain()
        {
            InitializeComponent();
            LoadCPCFiles();
            xgvOfficialFile.BestFitColumns(true);
        }

        public Hashtable DicOfficialCodeName()
        {
            var ht = new Hashtable
            {
                {"250340", "pe_pass"},
                {"210304", "pe_pass"},
                {"200028", "procedure_accepted"},
                {"250302", "filing_receipt"},
                {"200021", "filing_receipt"},
                {"210302", "corr_other"}
            };
            return ht;
        }

        public DataTable GetFiles()
        {
            return DbHelperMySQL.Query("SELECT * FROM dzsq_med_tzs").Tables[0];
        }

        public DataTable ExistCase(string sAppNo, string sOurNo)
        {
            var dt = new DataTable();
            if (!string.IsNullOrWhiteSpace(sAppNo))
            {
                if (!sAppNo.Contains('.'))
                    sAppNo = sAppNo.Insert(sAppNo.Length - 1, ".");
                dt = DbHelperOra.Query($"select OURNO,APPLICATION_NO from PATENTCASE where APPLICATION_NO = '{sAppNo}'").Tables[0];
            }
            if (!string.IsNullOrWhiteSpace(sOurNo) && dt.Rows.Count == 0)
            {
                return DbHelperOra.Query($"select * from PATENTCASE where OURNO like '{sOurNo}%'").Tables[0];
            }
            return dt;
        }

        public bool ExistFile(string sOurNo, string sAppNo, string sFileCode)
        {
            var sContent = DicOfficialCodeName()[sFileCode];
            return DbHelperOra.Exists($"select 1 from RECEIVINGLOG where (OURNO = '{sOurNo}' or APPNO = '{sAppNo}') and CONTENT = '{sContent}'");
        }

        public void LoadCPCFiles()
        {
            try
            {
                SplashScreenManager.ShowDefaultWaitForm();
                var listFiles = GetFiles().Rows.Cast<DataRow>().Select(r => new CPCOfficialFile(r)).ToList();
                xgcOfficialFile.DataSource = listFiles;
                foreach (var cpcOffcialFile in listFiles)
                {
                    Application.DoEvents();
                    var dtCase = ExistCase(cpcOffcialFile.AppNo, cpcOffcialFile.CPCSerial);
                    if (dtCase.Rows.Count < 1)
                    {
                        cpcOffcialFile.Note += "未找到案件";
                        continue;
                    }
                    if (dtCase.Rows.Count > 1)
                    {
                        cpcOffcialFile.Note += "找到多个案件";
                        continue;
                    }
                    cpcOffcialFile.CaseSerial = dtCase.Rows[0]["OURNO"].ToString();
                    cpcOffcialFile.AppNo = dtCase.Rows[0]["APPLICATION_NO"].ToString();
                    if (ExistFile(cpcOffcialFile.CaseSerial, cpcOffcialFile.AppNo, cpcOffcialFile.FileCode))
                    {
                        cpcOffcialFile.Note += "通知书已在系统中存在";
                        continue;
                    }
                    xgcOfficialFile.Refresh();
                }
            }
            catch (Exception exception)
            {
                XtraMessageBox.Show(exception.ToString());
            }
            finally
            {
                SplashScreenManager.CloseDefaultWaitForm();
            }
        }

        public void GeneratePDFFile(CPCOfficialFile cpcOfficialFile)
        {
            if (!Directory.Exists(cpcOfficialFile.FilePath)) return;
            var sCPCFilePath = $@"{cpcOfficialFile.FilePath}\{cpcOfficialFile.FileSerial}\{cpcOfficialFile.FileSerial}";
            if (!Directory.Exists(sCPCFilePath)) return;
            var files = Directory.GetFiles(sCPCFilePath, "*.tif").ToList();
            Directory.GetFiles(cpcOfficialFile.FilePath, "*.tif", SearchOption.AllDirectories).ToList().ForEach(s =>
            {
                if (files.Contains(s)) return;
                files.Add(s);
            });
            if (files.Count == 0) return;
            var sGenerateFile = $@"D:\GenerateFolder\{cpcOfficialFile.FileSerial}.pdf";

            try
            {
                var document = new Document(PageSize.A4, 25, 25, 25, 25);
                if (!Directory.Exists(Path.GetDirectoryName(sGenerateFile)))
                    Directory.CreateDirectory(Path.GetDirectoryName(sGenerateFile));
                PdfWriter.GetInstance(document, new FileStream(sGenerateFile, FileMode.Create));
                document.Open();
                foreach (var p in files)
                {
                    var image = iTextSharp.text.Image.GetInstance(p);
                    if (image.Height > PageSize.A4.Height - 25)
                    {
                        image.ScaleToFit(PageSize.A4.Width - 25, PageSize.A4.Height - 25);
                    }
                    else if (image.Width > PageSize.A4.Width - 25)
                    {
                        image.ScaleToFit(PageSize.A4.Width - 25, PageSize.A4.Height - 25);
                    }
                    image.Alignment = Element.ALIGN_MIDDLE;
                    document.NewPage();
                    document.Add(image);
                }
                document.Close();
            }
            catch (Exception e)
            {
                cpcOfficialFile.Note = e.ToString();
            }
            cpcOfficialFile.BizFilePath = sGenerateFile;
        }

        private void xbbiImport_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            var listCPCOfficialFile = xgvOfficialFile.GetSelectedRows().Select(h => xgvOfficialFile.GetRow(h) as CPCOfficialFile).Where(f => f != null).ToList();
            listCPCOfficialFile.ForEach(f =>
            {
                GeneratePDFFile(f);
            });
            xgvOfficialFile.RefreshData();
        }

        private void xgvOfficialFile_DoubleClick(object sender, EventArgs e)
        {
            var cpcFile = xgvOfficialFile.GetFocusedRow() as CPCOfficialFile;
            if (cpcFile == null || string.IsNullOrWhiteSpace(cpcFile.BizFilePath)) return;
            Process.Start(cpcFile.BizFilePath);
        }

        private void xbbiRefresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadCPCFiles();
            xgvOfficialFile.RefreshData();
        }
    }
}
