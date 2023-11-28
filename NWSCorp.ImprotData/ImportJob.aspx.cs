using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Sitecore.Data;
using Sitecore.Data.Fields;
using Sitecore.Data.Items;
using Sitecore.Globalization;
using Sitecore.SecurityModel;

namespace NWSCorp.ImprotData
{
    public partial class ImportJob : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Page.DataBind();
        }

        public string Logstring { get; set; }
        protected void NextButton_Click(object sender, EventArgs e)
        {
            try
            {
                LoadExcel(ImportData);
            }
            catch (Exception ex)
            {
                Label1.Text = "{ERROR:" + ex.Message + "}";
            }
        }

        private void LoadExcel(FileUpload file)
        {
            using (MemoryStream stream = new MemoryStream(file.FileBytes))
            {
                IWorkbook excel = WorkbookFactory.Create(stream);
                if (excel == null)
                {
                    return;
                }

                var sheetTable = excel.GetSheetAt(0);

                var joTable = BuildTable(sheetTable);

                Language language = Language.Parse("en");
                if (RadioButton2.Checked)
                {
                    language = Language.Parse("zh-hk");
                }
                else if (RadioButton3.Checked)
                {
                    language = Language.Parse("zh-cn");
                }

                CreateNewItem(joTable, language);
            }
        }


        private DataTable BuildTable(ISheet sheetTable)
        {
            DataTable dt = new DataTable();
            IRow headrow = sheetTable.GetRow(0);

            int headrowCount = headrow.Cells.Where(s => !string.IsNullOrWhiteSpace(s.StringCellValue)).Count();

            for (int i = headrow.FirstCellNum; i < headrowCount; i++)
            {
                DataColumn datacolum = new DataColumn(headrow.Cells[i].StringCellValue.Trim());
                dt.Columns.Add(datacolum);
            }

            for (int r = 1; r <= sheetTable.LastRowNum; r++)
            {
                DataRow dr = dt.NewRow();

                IRow row = sheetTable.GetRow(r);

                for (int j = 0; j < headrowCount; j++)
                {
                    if (row != null)
                    {
                        ICell cell = row.GetCell(j);
                        dr[j] = GetCellValue(cell).Trim();
                    }
                }
                dt.Rows.Add(dr);

            }
            return dt;
        }

        private static string GetCellValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank: //空数据类型 
                    return string.Empty;
                case CellType.Boolean: //bool类型
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.String: //string 类型
                    return cell.StringCellValue;
                case CellType.Formula: //带公式类型
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.String:
                            var formulaValue = cell.StringCellValue;
                            return formulaValue;

                        case CellType.Numeric:
                            var formulaNumericValue = cell.NumericCellValue;
                            return formulaNumericValue.ToString();

                        case CellType.Boolean:
                            var formulaBoolValue = cell.BooleanCellValue;
                            return formulaBoolValue.ToString();
                        default:
                            return string.Empty;
                    }
                case CellType.Numeric: //数字类型
                    if (HSSFDateUtil.IsCellDateFormatted(cell) &&
                        (cell.CellStyle.DataFormat == 0x14 || cell.CellStyle.DataFormat == 0x15))//日期类型
                    {
                        var dt = GetJavaCalendar(cell.NumericCellValue, false, null, false);
                        if (dt != null)
                            return dt.ToString("HH:mm:ss");
                        else
                            return string.Empty;
                    }
                    else if (HSSFDateUtil.IsCellDateFormatted(cell))
                    {
                        var dt = GetJavaCalendar(cell.NumericCellValue, false, null, false);
                        if (dt != null)
                            //return dt.ToString("dd/MM/yyyy");
                            return dt.ToString("yyyyMMdd")+"T"+dt.ToString("HHmmss")+"Z";
                        else
                            return string.Empty;
                    }
                    else //其它数字
                    {
                        return cell.NumericCellValue.ToString();
                    }
                case CellType.Unknown: //无法识别类型
                default: //默认类型
                    return cell.ToString();//


            }
        }

        public static DateTime GetJavaCalendar(double date, bool use1904windowing, TimeZone timeZone, bool roundSeconds)
        {
            int num = (int)Math.Floor(date);
            int millisecondsInDay = (int)((date - (double)num) * 86400000.0 + 0.5);
            return SetCalendar(num, millisecondsInDay, use1904windowing, roundSeconds);
        }

        public static DateTime SetCalendar(int wholeDays, int millisecondsInDay, bool use1904windowing, bool roundSeconds)
        {
            int year = 1900;
            int num = -1;
            if (use1904windowing)
            {
                year = 1904;
                num = 1;
            }
            else if (wholeDays < 61)
            {
                num = 0;
            }

            DateTime dateTime = new DateTime(year, 1, 1).AddDays(wholeDays + num - 1).AddMilliseconds(millisecondsInDay);

            return dateTime;
        }


        private void CreateNewItem(DataTable dt, Language language)
        {
            try
            {
                int count = 0;

                Database masterDb = Sitecore.Configuration.Factory.GetDatabase("master");

                Item opportunitiesItem = masterDb.GetItem(new ID("{CFEE8706-2B4E-4E2C-89E4-736F66DD0C30}"));
                TemplateItem jobTemplate = masterDb.GetTemplate(new ID("{9CA83F3E-458F-462C-8DCF-C9E612157523}"));
                var parentItemList = opportunitiesItem.Children;

                foreach (var dr in dt.Rows)
                {
                    var cName = (dr as DataRow)["Company Name"].ToString();

                    if (!string.IsNullOrWhiteSpace(cName))
                    {
                        var parentItem = parentItemList.FirstOrDefault(x => x.Name == cName);
                        if (parentItem != null)
                        {
                            var jobItemList = parentItem.Children;
                            var jobItem = jobItemList.FirstOrDefault(x => x.Name == (dr as DataRow)["Position"].ToString());
                            if (jobItem == null)
                            {
                                using (new SecurityDisabler())
                                {
                                    // must en
                                    Item newItem = parentItem.Add(ItemUtil.ProposeValidItemName((dr as DataRow)["Position"].ToString()), jobTemplate);
                                    //newItem.Versions.RemoveAll(false);
                                    Item newItem_l = masterDb.GetItem(newItem.ID,language);
                                    if (newItem_l.IsFallback) {
                                        newItem_l = newItem_l.Versions.AddVersion();
                                    }
                                    //

                                    newItem_l.Editing.BeginEdit();
                                    try
                                    {
                                        newItem_l["Position"] = (dr as DataRow)["Position"].ToString();

                                        var funid = GetOptionId("Functions", (dr as DataRow)["Functions"].ToString());
                                        newItem_l["Functions"] = funid;

                                        var empid = GetOptionId("Employment type", (dr as DataRow)["Employment type"].ToString());
                                        newItem_l["Employment type"] = empid;

                                        var locid = GetOptionId("Location", (dr as DataRow)["Location"].ToString());
                                        newItem_l["Location"] = locid;

                                        newItem_l["Job Details"] = (dr as DataRow)["Job Details"].ToString();
                                        newItem_l["Job Requirements"] = (dr as DataRow)["Job Requirements"].ToString();

                                        var indid = GetOptionId("Industries", (dr as DataRow)["Industries"].ToString());
                                        newItem_l["Industries"] = indid;

                                        var worid = GetOptionId("Workplace Type", (dr as DataRow)["Workplace Type"].ToString());
                                        newItem_l["Workplace Type"] = worid;

                                        var carid = GetOptionId("Career Level", (dr as DataRow)["Career Level"].ToString());
                                        newItem_l["Career Level"] = carid;

                                        var eduid = GetOptionId("Education level", (dr as DataRow)["Education level"].ToString());
                                        newItem_l["Education level"] = eduid;

                                        var benid = GetOptionId("Benefits", (dr as DataRow)["Benefits"].ToString());
                                        newItem_l["Benefits"] = benid;

                                        newItem_l["Post date"] = (dr as DataRow)["Post date"].ToString();
                                        newItem_l["SAP Job application link"] = ChangeLink((dr as DataRow)["SAP Job application link"].ToString());
                                        newItem_l["External Job application link"] = ChangeLink((dr as DataRow)["External Job application link"].ToString());

                                        var shaid = GetSharingId("Sharing Job", (dr as DataRow)["Sharing Job"].ToString());
                                        newItem_l["Sharing Job"] = shaid;

                                        newItem_l.Editing.EndEdit();
                                        count++;
                                        Logstring += "{'" + cName + "' succeed to add job:'" + (dr as DataRow)["Position"].ToString() + "'}";
                                    }
                                    catch (Exception ex)
                                    {
                                        newItem_l.Editing.CancelEdit();
                                        Logstring += "{" + cName + "-ERROR:" + ex.Message + "}";
                                        //throw;
                                    }
                                }
                            }
                            else
                            {
                                using (new SecurityDisabler())
                                {
                                    Item jobItem_l = masterDb.GetItem(jobItem.ID, language);
                                    if (jobItem_l.IsFallback)
                                    {
                                        jobItem_l = jobItem_l.Versions.AddVersion();
                                    }

                                    jobItem_l.Editing.BeginEdit();
                                    try
                                    {
                                        jobItem_l["Position"] = (dr as DataRow)["Position"].ToString();

                                        var funid = GetOptionId("Functions", (dr as DataRow)["Functions"].ToString());
                                        jobItem_l["Functions"] = funid;

                                        var empid = GetOptionId("Employment type", (dr as DataRow)["Employment type"].ToString());
                                        jobItem_l["Employment type"] = empid;

                                        var locid = GetOptionId("Location", (dr as DataRow)["Location"].ToString());
                                        jobItem_l["Location"] = locid;

                                        jobItem_l["Job Details"] = (dr as DataRow)["Job Details"].ToString();
                                        jobItem_l["Job Requirements"] = (dr as DataRow)["Job Requirements"].ToString();

                                        var indid = GetOptionId("Industries", (dr as DataRow)["Industries"].ToString());
                                        jobItem_l["Industries"] = indid;

                                        var worid = GetOptionId("Workplace Type", (dr as DataRow)["Workplace Type"].ToString());
                                        jobItem_l["Workplace Type"] = worid;

                                        var carid = GetOptionId("Career Level", (dr as DataRow)["Career Level"].ToString());
                                        jobItem_l["Career Level"] = carid;

                                        var eduid = GetOptionId("Education level", (dr as DataRow)["Education level"].ToString());
                                        jobItem_l["Education level"] = eduid;

                                        var benid = GetOptionId("Benefits", (dr as DataRow)["Benefits"].ToString());
                                        jobItem_l["Benefits"] = benid;

                                        jobItem_l["Post date"] = (dr as DataRow)["Post date"].ToString();
                                        jobItem_l["SAP Job application link"] = ChangeLink((dr as DataRow)["SAP Job application link"].ToString());
                                        jobItem_l["External Job application link"] = ChangeLink((dr as DataRow)["External Job application link"].ToString());

                                        var shaid = GetSharingId("Sharing Job", (dr as DataRow)["Sharing Job"].ToString());
                                        jobItem_l["Sharing Job"] = shaid;

                                        jobItem_l.Editing.EndEdit();
                                        count++;
                                        Logstring += "{'" + cName + "' succeed to update job:'" + (dr as DataRow)["Position"].ToString() + "'}";
                                    }
                                    catch (Exception ex)
                                    {
                                        jobItem_l.Editing.CancelEdit();
                                        Logstring += "{" + cName + "-ERROR:" + ex.Message + "}";
                                        //throw;
                                    }
                                }
                                //Logstring += "{'" + (dr as DataRow)["Position"].ToString() + "' had exist in '" + cName + "' }";
                            }
                        }
                        else //company not exist
                        {
                        }
                    }
                }
                Logstring += " { Succeed to add/update " + count + " job(s) }";
            }
            catch (Exception ex)
            {
                //throw;
            }
        }
        private string GetOptionId(string fieldsName, string optionText)
        {
            var optionList = optionText.Split(',').ToList();    //Excel 多选用","分隔？
            Database masterDb = Sitecore.Configuration.Factory.GetDatabase("master");
            TemplateItem jobTemplate = masterDb.GetTemplate(new ID("{9CA83F3E-458F-462C-8DCF-C9E612157523}"));
            var optinsId = jobTemplate.StandardValues.Fields[fieldsName].Source;
            var funcitonsItem = masterDb.GetItem(new ID(optinsId));
            var options = funcitonsItem.Children.Where(x => x.TemplateID.ToString() == "{479F03FF-3A04-4AB1-9B29-F0B576BFC643}").Where(x => optionList.Contains(x.Fields["Text"].Value)).Select(x => x.ID.ToString()).ToList();
            if (options.Count > 0)
            {
                return string.Join("|", options);
            }
            return string.Empty;
        }
        private string GetSharingId(string fieldsName, string optionText)
        {
            var optionList = optionText.Split(',').ToList();    //Excel 多选用","分隔？
            Database masterDb = Sitecore.Configuration.Factory.GetDatabase("master");
            TemplateItem jobTemplate = masterDb.GetTemplate(new ID("{9CA83F3E-458F-462C-8DCF-C9E612157523}"));
            var optinsId = jobTemplate.StandardValues.Fields[fieldsName].Source;
            var funcitonsItem = masterDb.GetItem(new ID(optinsId));
            var options = funcitonsItem.Children.Where(x => x.TemplateID.ToString() == "{8297CA9A-F6A3-4A06-BF92-6A4E8F6EB92F}").Where(x => optionList.Contains(x.Name)).Select(x => x.ID.ToString()).ToList();
            if (options.Count > 0)
            {
                return string.Join("|", options);
            }
            return string.Empty;
        }

        private string ChangeLink(string link)
        {
            //< link linktype = "external" url = "www.google.com" anchor = "" target = "" />
            return "<link linktype=\"external\" url=\""+link+"\" anchor=\"\" target=\"\" />";
        }
    }

}