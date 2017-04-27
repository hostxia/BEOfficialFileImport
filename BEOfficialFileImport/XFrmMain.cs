using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

        public DataTable GetFiles()
        {
            return DbHelperMySQL.Query("SELECT * FROM dzsq_med_tzs").Tables[0];
        }

        public DataTable ExistCase(string sAppNo, string sOurNo)
        {
            var dt = new DataTable();
            if (!string.IsNullOrWhiteSpace(sAppNo))
            {
                dt = DbHelperOra.Query($"select OURNO,APPLICATION_NO,CLIENT_NUMBER,DOCSTATE from PATENTCASE where APPLICATION_NO = '{sAppNo}'").Tables[0];
            }
            if (!string.IsNullOrWhiteSpace(sOurNo) && dt.Rows.Count == 0)
            {
                return DbHelperOra.Query($"select OURNO,APPLICATION_NO,CLIENT_NUMBER,DOCSTATE from PATENTCASE where OURNO like '{sOurNo}%'").Tables[0];
            }
            return dt;
        }

        public bool ExistFile(string sOurNo, string sAppNo, string sFileName, DateTime dtSendDate, string sFileCode = "")
        {
            return DbHelperOra.Exists($"select 1 from RECEIVINGLOG where (OURNO = '{sOurNo}' or APPNO = '{sAppNo}') and COMMENTS like '%{sFileName}%' and ISSUEDATE = to_date('{dtSendDate:yyyy/MM/dd}','yyyy/MM/dd')");
        }

        public void LoadCPCFiles()
        {
            var listFiles = GetFiles().Rows.Cast<DataRow>().Select(r => new CPCOfficialFile(r)).ToList();
            xgcOfficialFile.DataSource = listFiles;
            xgcOfficialFile.Refresh();
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
            var sGenerateFile = $@"D:\GenerateFolder\{DateTime.Now:yyyyMMdd}\{cpcOfficialFile.CaseSerial.Substring(0, cpcOfficialFile.CaseSerial.IndexOf("-"))}-{DateTime.Now:yyyyMMdd}-{(cpcOfficialFile.CPCOfficialFileConfig == null ? cpcOfficialFile.FileName : cpcOfficialFile.CPCOfficialFileConfig.Rename)}.pdf";

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
            SplashScreenManager.ShowDefaultWaitForm();
            var listCPCOfficialFile =
                xgvOfficialFile.GetSelectedRows()
                    .Select(h => xgvOfficialFile.GetRow(h) as CPCOfficialFile)
                    .Where(f => f != null)
                    .ToList();

            listCPCOfficialFile.ForEach(f =>
            {
                var dtNow = DateTime.Now;
                f.Note = string.Empty;
                Application.DoEvents();
                try
                {
                    if (f.SendDate == DateTime.MinValue)
                    {
                        f.Note += "官方发文日为空，该官文无法导入";
                        return;
                    }
                    if (!string.IsNullOrWhiteSpace(f.AppNo) && !f.AppNo.Contains('.'))
                        f.AppNo = f.AppNo.Insert(f.AppNo.Length - 1, ".");
                    var dtCase = ExistCase(f.AppNo, f.CPCSerial);
                    if (dtCase.Rows.Count < 1)
                    {
                        f.Note += "未找到案件";
                        return;
                    }
                    if (dtCase.Rows.Count > 1)
                    {
                        f.Note += "找到多个案件";
                        return;
                    }
                    f.CaseSerial = dtCase.Rows[0]["OURNO"].ToString();
                    GeneratePDFFile(f);
                    if (ExistFile(f.CaseSerial, f.AppNo, f.FileName, f.SendDate))
                    {
                        f.Note += "通知书已在系统中存在";
                        return;
                    }

                    var strSql = new List<string>();
                    if (!string.IsNullOrWhiteSpace(f.AppNo))
                        strSql.Add($"update PATENTCASE set APPLICATION_NO = '{f.AppNo}' where OURNO = '{f.CaseSerial}'");
                    else
                        f.AppNo = dtCase.Rows[0][1].ToString();


                    if (ExistFile(f.CaseSerial, f.AppNo, "国际申请进入中国国家阶段通知书", f.SendDate) || listCPCOfficialFile.Any(a => a.FileCode == "250302" && a.SendDate.Date == f.SendDate.Date && a.AppNo == f.AppNo))//存在当天的国际申请进入中国国家阶段通知书
                    {
                        f.CPCOfficialFileConfig.Dealer = "DXD";
                    }

                    if (dtCase.Rows[0][3]?.ToString().Trim().ToUpper() == "T")
                    {
                        f.CPCOfficialFileConfig.Dealer = "GS";
                    }
                    strSql.Add($"insert into RECEIVINGLOG (PID,ISSUEDATE,RECEIVED,SENDERID,SENDER,OURNO,APPNO,CLIENTNO,CONTENT,COPIES,COMMENTS,STATUS,HANDLER) values ('{DateTime.Now:yyyyMMdd_HHmmss_ffffff_0}',to_date('{f.SendDate.Date:yyyy/MM/dd}','yyyy/MM/dd'),to_date('{DateTime.Now.Date:yyyy/MM/dd}','yyyy/MM/dd'),'SIPO','SIPO','{dtCase.Rows[0][0]}','{dtCase.Rows[0][1]}','{dtCase.Rows[0][2]}','other','1','{f.FileName}','P','{f.CPCOfficialFileConfig?.Dealer}')");

                    if (f.CPCOfficialFileConfig?.DeadlineFiledType != null)
                    {
                        switch (f.CPCOfficialFileConfig.DeadlineFiledType.Value)
                        {
                            case DeadlineFiledType.Case:
                                if (f.CPCOfficialFileConfig.DeadlineFiled == "GRANTNOTIC_DATE")
                                    strSql.Add($"update PATENTCASE set GRANTNOTIC_DATE=to_date('{f.SendDate.Date:yyyy/MM/dd}','yyyy/MM/dd'),REGFEE_DL=to_date('{f.SendDate.Date.AddDays(15).AddMonths(2):yyyy/MM/dd}','yyyy/MM/dd') where OURNO = '{f.CaseSerial}'");//更新办登信息
                                else if (f.CPCOfficialFileConfig.DeadlineFiled == "PRE_EXAM_PASSED")
                                    strSql.Add($"update PATENTCASE set PRE_EXAM_PASSED=to_date('{f.SendDate.Date:yyyy/MM/dd}','yyyy/MM/dd') where OURNO = '{f.CaseSerial}'");//更新初审合格日
                                break;
                            case DeadlineFiledType.OA:
                                strSql.Add(
                                    $"insert into GENERALALERT (CREATED,TYPEID,OURNO,TRIGERDATE1,DUEDATE,OATYPE,COMMENTS) values (to_date('{DateTime.Now:yyyy/MM/dd HH:mm:ss}','yyyy/MM/dd hh24:mi:ss'),'invoa','{f.CaseSerial}',to_date('{f.SendDate.Date:yyyy/MM/dd}','yyyy/MM/dd'),to_date('{f.SendDate.Date.AddDays(f.CPCOfficialFileConfig.AddDays).AddMonths(f.CPCOfficialFileConfig.AddMonths):yyyy/MM/dd}','yyyy/MM/dd'),'{f.CPCOfficialFileConfig.DeadlineFiled}','{f.CPCOfficialFileConfig.DeadlineFiledNote}')");
                                break;
                            case DeadlineFiledType.Deadline:
                                strSql.Add(
                                    $"insert into GENERALALERT (CREATED,TYPEID,OURNO,DUEDATE,COMMENTS) values (to_date('{DateTime.Now:yyyy/MM/dd HH:mm:ss}','yyyy/MM/dd hh24:mi:ss'),'{f.CPCOfficialFileConfig.DeadlineFiled}','{f.CaseSerial}',to_date('{f.SendDate.Date.AddDays(f.CPCOfficialFileConfig.AddDays).AddMonths(f.CPCOfficialFileConfig.AddMonths):yyyy/MM/dd}','yyyy/MM/dd'),'{f.CPCOfficialFileConfig.DeadlineFiledNote}')");
                                break;
                            case DeadlineFiledType.FCaseDeadline:
                                break;
                        }
                    }

                    var array = new ArrayList();
                    array.AddRange(strSql);
                    DbHelperOra.ExecuteSqlTran(array);
                    f.Note = "已导入";
                }
                catch (Exception exception)
                {
                    f.Note += exception.ToString();
                }
                while (dtNow.AddSeconds(1) > DateTime.Now)
                {

                }

            });
            xgvOfficialFile.RefreshData();
            SplashScreenManager.CloseDefaultWaitForm();
        }

        private void xgvOfficialFile_DoubleClick(object sender, EventArgs e)
        {
            var cpcFile = xgvOfficialFile.GetFocusedRow() as CPCOfficialFile;
            if (string.IsNullOrWhiteSpace(cpcFile?.BizFilePath)) return;
            Process.Start(cpcFile.BizFilePath);
        }

        private void xbbiRefresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadCPCFiles();
            xgvOfficialFile.RefreshData();
        }

        private void xbbiDelete_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            var listCPCOfficialFile =
                xgvOfficialFile.GetSelectedRows()
                    .Select(h => xgvOfficialFile.GetRow(h) as CPCOfficialFile)
                    .Where(f => f != null)
                    .ToList();

            var listCPCOfficialFiles = xgcOfficialFile.DataSource as List<CPCOfficialFile>;
            listCPCOfficialFile.ForEach(c => listCPCOfficialFiles.Remove(c));
            xgcOfficialFile.RefreshDataSource();
        }

        public void SetWangruiInfo()
        {

        }
    }
}
