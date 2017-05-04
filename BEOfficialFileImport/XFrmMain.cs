using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
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
            var dtNow = DateTime.Now;
            if (DateTime.Now < new DateTime(dtNow.Year, dtNow.Month, dtNow.Day, 12, 0, 0))
                return DbHelperMySQL.Query($"SELECT * FROM dzsq_med_tzs where XIAZAIRQ >= '{DateTime.Now.Date}' and XIAZAIRQ < '{new DateTime(dtNow.Year, dtNow.Month, dtNow.Day, 12, 0, 0)}'").Tables[0];
            return DbHelperMySQL.Query($"SELECT * FROM dzsq_med_tzs where XIAZAIRQ >= '{new DateTime(dtNow.Year, dtNow.Month, dtNow.Day, 12, 0, 0)}' and XIAZAIRQ < '{dtNow.Date.AddDays(1).AddSeconds(-1)}'").Tables[0];
        }

        public DataTable ExistCase(string sAppNo, string sOurNo)
        {
            var dt = new DataTable();
            if (!string.IsNullOrWhiteSpace(sAppNo))
            {
                dt = DbHelperOra.Query($"select OURNO,APPLICATION_NO,CLIENT_NUMBER,DOCSTATE,CLIENT,CLIENT_NAME,APPL_CODE1,APPLICANT1,APPLICANT_CH1,APPL_CODE2,APPLICANT2,APPLICANT_CH2,APPL_CODE3,APPLICANT3,APPLICANT_CH3,APPL_CODE4,APPLICANT4,APPLICANT_CH4,APPL_CODE5,APPLICANT5,APPLICANT_CH5 from PATENTCASE where APPLICATION_NO = '{sAppNo}'").Tables[0];
            }
            if (!string.IsNullOrWhiteSpace(sOurNo) && dt.Rows.Count == 0)
            {
                return DbHelperOra.Query($"select OURNO,APPLICATION_NO,CLIENT_NUMBER,DOCSTATE,CLIENT,CLIENT_NAME,APPL_CODE1,APPLICANT1,APPLICANT_CH1,APPL_CODE2,APPLICANT2,APPLICANT_CH2,APPL_CODE3,APPLICANT3,APPLICANT_CH3,APPL_CODE4,APPLICANT4,APPLICANT_CH4,APPL_CODE5,APPLICANT5,APPLICANT_CH5 from PATENTCASE where OURNO like '{sOurNo}%'").Tables[0];
            }
            return dt;
        }

        public bool ExistFile(string sOurNo, string sAppNo, string sFileName, DateTime dtSendDate, string sFileCode = "")
        {
            return DbHelperOra.Exists($"select 1 from RECEIVINGLOG where (OURNO = '{sOurNo}' or APPNO = '{sAppNo}') and COMMENTS like '%{sFileName}%' and ISSUEDATE = to_date('{dtSendDate:yyyy/MM/dd}','yyyy/MM/dd')");
        }

        public void LoadCPCFiles()
        {
            SplashScreenManager.ShowDefaultWaitForm();
            var listFiles = GetFiles().Rows.Cast<DataRow>().Select(r => new CPCOfficialFile(r)).ToList();
            listFiles.ForEach(f =>
            {
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
                    if (string.IsNullOrWhiteSpace(f.AppNo))
                        f.AppNo = dtCase.Rows[0][1].ToString();

                    if (dtCase.Rows[0][3]?.ToString().Trim().ToUpper() == "T")
                    {
                        f.CPCOfficialFileConfig.Dealer = "GS";
                    }
                    f.CaseSerial = dtCase.Rows[0]["OURNO"].ToString();
                    f.ClientNo = dtCase.Rows[0]["CLIENT"].ToString();
                    f.ClientName = dtCase.Rows[0]["CLIENT_NAME"].ToString();

                    f.CPCOfficialFileConfig.Dealer = HandlerRedistribution(f.CaseSerial, f.CPCOfficialFileConfig.Dealer);
                    GeneratePDFFile(f);
                }
                catch (Exception exception)
                {
                    f.Note += exception.ToString();
                }
            });
            xgcOfficialFile.DataSource = listFiles;
            xgcOfficialFile.Refresh();
            SplashScreenManager.CloseDefaultWaitForm();
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
            var sGenerateFile = $@"D:\GenerateFolder\{DateTime.Now:yyyyMMddtt}\{cpcOfficialFile.CaseSerial.Substring(0, cpcOfficialFile.CaseSerial.IndexOf("-", StringComparison.Ordinal))}-{DateTime.Now:yyyyMMdd}-{(cpcOfficialFile.CPCOfficialFileConfig == null ? cpcOfficialFile.FileName : cpcOfficialFile.CPCOfficialFileConfig.Rename)}.pdf";

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
                    if (string.IsNullOrWhiteSpace(f.CPCOfficialFileConfig.Dealer))
                    {
                        f.Note += "请填写处理人";
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
                    f.ClientNo = dtCase.Rows[0]["CLIENT"].ToString();
                    f.ClientName = dtCase.Rows[0]["CLIENT_NAME"].ToString();

                    f.Applicants = new Hashtable();
                    for (int i = 1; i <= 5; i++)
                    {
                        if (string.IsNullOrWhiteSpace(dtCase.Rows[0][$"APPL_CODE{i}"].ToString())) continue;
                        f.Applicants.Add(dtCase.Rows[0][$"APPL_CODE{i}"].ToString(), !string.IsNullOrWhiteSpace(dtCase.Rows[0][$"APPLICANT{i}"].ToString()) ? dtCase.Rows[0][$"APPLICANT{i}"].ToString() : dtCase.Rows[0][$"APPLICANT_CH{i}"].ToString());
                    }
                    if (ExistFile(f.CaseSerial, f.AppNo, f.FileName, f.SendDate))
                    {
                        f.Note += "通知书已在系统中存在";
                        return;
                    }
                    var strSql = new List<string>();
                    if (!string.IsNullOrWhiteSpace(f.AppNo))
                        strSql.Add($"update PATENTCASE set APPLICATION_NO = '{f.AppNo}' where OURNO = '{f.CaseSerial}'");

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
                                if (f.FileCode == "200702" && !IsValidCase(f.AppNo))//如果是专利权终止通知书且案件已届满
                                    break;
                                strSql.Add(
                                    $"insert into GENERALALERT (CREATED,TYPEID,OURNO,DUEDATE,COMMENTS) values (to_date('{DateTime.Now:yyyy/MM/dd HH:mm:ss}','yyyy/MM/dd hh24:mi:ss'),'{f.CPCOfficialFileConfig.DeadlineFiled}','{f.CaseSerial}',to_date('{f.SendDate.Date.AddDays(f.CPCOfficialFileConfig.AddDays).AddMonths(f.CPCOfficialFileConfig.AddMonths):yyyy/MM/dd}','yyyy/MM/dd'),'{f.CPCOfficialFileConfig.DeadlineFiledNote}')");
                                break;
                            case DeadlineFiledType.FCaseDeadline:
                                break;
                        }
                    }

                    SendEmail(f);
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

        private bool IsValidCase(string sAppNo)
        {
            var objDate = DbHelperOra.GetSingle($"select FILING_DATE from PATENTCASE where APPLICATION_NO = '{sAppNo}'");
            if (objDate == null) return true;
            DateTime dt;
            DateTime.TryParse(objDate.ToString(), out dt);
            if (dt == DateTime.MinValue) return true;
            var cType = sAppNo.Length > 10 ? sAppNo[4] : sAppNo[2];
            if (cType == '1' || cType == '8')
                return dt.AddYears(20) >= DateTime.Now;
            return dt.AddYears(10) >= DateTime.Now;
        }

        private string HandlerRedistribution(string sOurNo, string sDefaultHandler)
        {
            if (sDefaultHandler == "ON") return "ON";
            var sOurNoShort = sOurNo.Substring(0, sOurNo.IndexOf("-", StringComparison.Ordinal));
            var sOurFlow = Regex.Match(sOurNoShort, @"\d{4}").Value;
            var listOurNum = sOurFlow.Reverse().ToList();
            if (sDefaultHandler == "DXD")
            {
                foreach (var sNum in listOurNum)
                {
                    if ("139".Contains(sNum))
                        return "DXD";
                    if ("058".Contains(sNum))
                        return "ZD";
                    if ("247".Contains(sNum))
                        return "QSY";
                }
            }
            else if (sDefaultHandler == "XN")
            {
                if ("134".Contains(listOurNum[0]))
                    return "XN";
                if ("267".Contains(listOurNum[0]))
                    return "SJY";
                if ("059".Contains(listOurNum[0]))
                    return "ZX";
                if ("8".Contains(listOurNum[0]))
                    return "ZNQ";
            }
            else if (sDefaultHandler == "GS")
            {
                if ("13579".Contains(listOurNum[0]))
                    return "LJJ";
                if ("24680".Contains(listOurNum[0]))
                    return "WRX";
            }
            else if (sDefaultHandler == "SP")
            {
                if ("13579".Contains(listOurNum[0]))
                    return "ZM";
                if ("24680".Contains(listOurNum[0]))
                    return "TSN";
            }
            return string.Empty;
        }

        private void SendEmail(CPCOfficialFile cpcOfficialFile)
        {
            var message = new MailMessage();
            var fromAddr = new MailAddress("official_notice@beijingeastip.com");
            message.From = fromAddr;
            message.To.Add(HtEmails[cpcOfficialFile.CPCOfficialFileConfig.Dealer].ToString());
            message.CC.Add("official_notice@beijingeastip.com");
            message.Headers.Add("Disposition-Notification-To", "official_notice@beijingeastip.com");

            var listSubject = new List<string>();
            listSubject.Add(cpcOfficialFile.CaseSerial);
            if (cpcOfficialFile.Applicants.Count > 0)
                listSubject.Add(cpcOfficialFile.Applicants.Cast<DictionaryEntry>().ToList()[0].Key.ToString());
            listSubject.Add(cpcOfficialFile.FileName);
            if (cpcOfficialFile.CPCOfficialFileConfig.CreateDeadline &&
                cpcOfficialFile.CPCOfficialFileConfig.DeadlineFiled != "PRE_EXAM_PASSED" &&
                !(cpcOfficialFile.FileCode == "200702" && !IsValidCase(cpcOfficialFile.AppNo)))
                listSubject.Add(
                    cpcOfficialFile.SendDate.AddDays(cpcOfficialFile.CPCOfficialFileConfig.AddDays)
                        .AddMonths(cpcOfficialFile.CPCOfficialFileConfig.AddMonths)
                        .ToString("yyyy/MM/dd"));
            message.Subject = string.Join("；", listSubject);

            var listBody = new List<string>();
            listBody.Add(
                $"申请人：{string.Join("； ", cpcOfficialFile.Applicants.Cast<DictionaryEntry>().Select(a => a.Key + "，" + a.Value).ToList())}");
            listBody.Add($"委托人：{cpcOfficialFile.ClientNo}，{cpcOfficialFile.ClientName}");
            listBody.Add($@"文件地址：\\PTFILE\PATENT\Cases-CN\{cpcOfficialFile.CaseSerial.Substring(0, cpcOfficialFile.CaseSerial.IndexOf("-", StringComparison.Ordinal))}\From_Office");
            message.Body = string.Join("\r\n", listBody);

            var client = new SmtpClient("smtp.beijingeastip.com", 25);
            client.Credentials = new NetworkCredential("official_notice@beijingeastip.com", "O@notice");
            //client.EnableSsl = true;
            client.Send(message);
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

        private void xbbiMessage_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xgvOfficialFile.GetSelectedRows().Select(xgvOfficialFile.GetRow).Cast<CPCOfficialFile>().ToList().ForEach(SendEmail);
        }

        private Hashtable HtEmails => new Hashtable()
        {
            {"DXD", "xiaoduo.ding@beijingeastip.com"},
            {"ZD", "di.zhang@beijingeastip.com"},
            {"QSY", "siyang.qin@beijingeastip.com"},
            {"GS", "shuang.guo@beijingeastip.com"},
            {"LJJ", "jianjiao.lu@beijingeastip.com"},
            {"WRX", "runxiu.wu@beijingeastip.com"},
            {"ZM", "meng.zhang@beijingeastip.com"},
            {"TSN", "shengnan.tian@beijingeastip.com"},
            {"XN", "na.xin@beijingeastip.com"},
            {"SJY", "jingyu.shen@beijingeastip.com"},
            {"ZX", "xiao.zhang@beijingeastip.com"},
            {"ZNQ", "naiqi.zhang@beijingeastip.com"},
            {"ON", "official_notice@beijingeastip.com"}
        };
    }
}
