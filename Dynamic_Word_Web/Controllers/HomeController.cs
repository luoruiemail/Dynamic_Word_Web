using Dynamic_Word_Web.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using Xceed.Words.NET;

namespace Dynamic_Word_Web.Controllers
{
    public class HomeController : Controller
    {
        /// <summary>
        /// 替换占位符测试数据
        /// </summary>
        private Dictionary<string, object> _replacePatterns = new Dictionary<string, object>()
            {
                { "#Seller#", "张三" },
                { "#ContractNumber#", "dfdfdf-834343-3432" },
                {"#TargetContractNumber#","dfdfdf-834343-3432" },
                { "#SignedAt#", "金融城" },
                { "#Buyer#", "李四" },
                { "#SignedDate#", "2018-12-08" },
                { "#Total#", "100" },
                { "#TotalCn#", "壹佰" },
                { "#SellerAddress#", "华侨城" },
                { "#SellerCorporation#", "****" },
                { "#SellerAuthorizedPerson#", "****" },
                { "#SellerPhoneNumber#", "132****6666" },
                { "#SellerBank#", "成都银行" },
                { "#SellerAccoutNumber#", "64531234276587678" },
                { "#BuyerAddress#", "科技城" },
                { "#BuyerCorporation#", "*****" },
                { "#BuyerAuthorizedPerson#", "王五" },
                { "#BuyerPhoneNumber#", "028-98789099" },
                { "#BuyerBank#", "工商银行" },
                { "#BuyerAccoutNumber#", "8767876789897667" },
                { "#SellerTaxNumber#", "87678767rrtrtr89897667" },
                { "#BuyerTaxNumber#", "werwrer67434343" },
                { "#Rules#","****************暂无********************" },
                { "#TableData#", new List<TableItem>{
                 new TableItem {T_COL_1="1",T_COL_2="废品1",T_COL_3="5",T_COL_4="15",T_COL_5="100",T_COL_6="吨",T_COL_7="100",TableIndex=TableIndex.T_Row_Type_1,T_COL_9="zybm0001" },
                 new TableItem {T_COL_1="1",T_COL_2="废品2",T_COL_3="6",T_COL_4="18",T_COL_5="101",T_COL_6="吨",T_COL_7="100",TableIndex=TableIndex.T_Row_Type_2,T_COL_9="zybm0001"  },
                 new TableItem {T_COL_1="2",T_COL_2="废品3",T_COL_3="7",T_COL_4="15",T_COL_5="100",T_COL_6="吨",T_COL_7="100",TableIndex=TableIndex.T_Row_Type_2 ,T_COL_9="zybm0001" },
                 new TableItem {T_COL_1="3",T_COL_2="废品4",T_COL_3="8",T_COL_4="15",T_COL_5="100",T_COL_6="吨",T_COL_7="100",TableIndex=TableIndex.T_Row_Type_2,T_COL_9="zybm0001"  }}
                },
            };

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        /// <summary>
        /// 下载word文件
        /// </summary>
        /// <param name="tempName">传入模板名称</param>
        /// <returns></returns>
        public ActionResult DownWordFile(string tempName = "")
        {
            if (tempName == "购销合同补充协议")
            {
                _replacePatterns["#TableData#"] = new List<TableItem>{
                 new TableItem {T_COL_1="产品ooP",T_COL_2="个",T_COL_3="5",T_COL_4="15",T_COL_5="100" } };
            }
            if (string.IsNullOrEmpty(tempName)) tempName = "服务费合同";
            var tempFile = $"/Content/ContractTemplate/{tempName}.docx";
            var fileName = tempFile.Substring(tempFile.LastIndexOf('/') + 1);
            MemoryStream stream = new MemoryStream();
            using (var doc = DocX.Load(Server.MapPath(tempFile)))
            {
                ReplaceDocText(doc, _replacePatterns);
                ReplaceDocTableText(doc, _replacePatterns["#TableData#"] as List<TableItem>);
                doc.SaveAs(stream);
                stream.Position = 0;
                byte[] tAryByte = stream.ToArray();
                stream.Close();
                stream.Dispose();
                return File(tAryByte,
                        //"application/msword",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        $"{DateTime.Now.ToString("yyyyMMddHHmmssffff") + fileName}{".docx"}");
            }
        }
        /// <summary>
        /// 替换单一占位符
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="_replacePatterns"></param>
        private void ReplaceDocText(DocX doc, Dictionary<string, object> _replacePatterns)
        {
            //处理多余的占位符
            Regex reg = new Regex("#.*?#");
            MatchCollection result = reg.Matches(doc.Text);
            foreach (var match in result)
            {
                var replaceText = match.ToString();
                if (_replacePatterns.Keys.Contains(replaceText))
                {
                    doc.ReplaceText(replaceText, _replacePatterns[replaceText].ToString());
                }
            }
        }
        /// <summary>
        /// 替换表格占位符数据
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="_replacePatterns"></param>
        private void ReplaceDocTableText(DocX doc, List<TableItem> tableData)
        {
            var table = doc.Tables[1];
            int cellCount = 0, index = 0, rowNum = 0;
            var tableRowList = table.Rows;
            foreach (var row in tableRowList)
            {
                if (index == 0) cellCount = row.Cells.Count;
                index++;
                //判断模板中表格行是否满足占位符条件，便于插入数据
                if (!new Regex("#t_col_.*?#").IsMatch(row.Cells[0].Paragraphs[0].Text)) continue;
                rowNum++;
                var tableResultList = ReturnList(tableData, rowNum);
                foreach (var item in tableResultList)
                {
                    //满足条件位置插入新行
                    var newRow = table.InsertRow(rowNum == 1 ? index - 1 : index);
                    for (int i = 0; i < cellCount; i++)
                    {
                        //记录新行单元格与占位符行单元格
                        Cell newCell = newRow.Cells[i], currentCell = row.Cells[i];
                        if (new Regex("#t_col_.*?#").IsMatch(currentCell.Paragraphs[0].Text))
                        {
                            newCell.VerticalAlignment = VerticalAlignment.Center;//垂直对其
                            newCell.Paragraphs[0].Alignment = Alignment.center;//水平对其                            
                            newCell.Paragraphs[0].Font("仿宋").FontSize(12);
                            UpdateTableCellValue(newCell.Paragraphs[0], currentCell.Paragraphs[0], item);
                        }
                    }
                }
            }
            DeleteRow(table);
        }

        /// <summary>
        /// 更新cell 值
        /// </summary>
        /// <param name="paragraph">新行cell</param>
        /// <param name="oldParagraph">占位符行cell</param>
        /// <param name="item"></param>
        private void UpdateTableCellValue(Paragraph paragraph, Paragraph oldParagraph, TableItem item)
        {
            switch (oldParagraph.Text)
            {
                case "#t_col_1#":
                    paragraph.Append(item.T_COL_1);
                    break;
                case "#t_col_2#":
                    paragraph.Append(item.T_COL_2);
                    break;
                case "#t_col_3#":
                    paragraph.Append(item.T_COL_3);
                    break;
                case "#t_col_4#":
                    paragraph.Append(item.T_COL_4);
                    break;
                case "#t_col_5#":
                    paragraph.Append(item.T_COL_5);
                    break;
                case "#t_col_6#":
                    paragraph.Append(item.T_COL_6);
                    break;
                case "#t_col_7#":
                    paragraph.Append(item.T_COL_7);
                    break;
                case "#t_col_8#":
                    paragraph.Append(item.T_COL_8);
                    break;
                case "#t_col_9#":
                    paragraph.Append(item.T_COL_9);
                    break;
                case "#t_col_10#":
                    paragraph.Append(item.T_COL_10);
                    break;
            }
        }

        /// <summary>
        /// 删除占位符行
        /// </summary>
        /// <param name="table"></param>
        private void DeleteRow(Table table)
        {
            var index = 0;
            var tableRows = table.Rows;
            foreach (var row in tableRows)
            {
                index++;
                var number = index - 1;
                if (new Regex("#t_col_.*?#").IsMatch(row.Cells[0].Paragraphs[0].Text))
                {
                    table.RemoveRow(number);
                    index--;
                }
            }
        }

        private IEnumerable<TableItem> ReturnList(List<TableItem> tableData, int rowNum = 1)
        {
            if (rowNum == 1) return tableData.Where(a => a.TableIndex == TableIndex.T_Row_Type_1).OrderByDescending(a => a.T_COL_1);
            if (rowNum == 2) return tableData.Where(a => a.TableIndex == TableIndex.T_Row_Type_2).OrderByDescending(a => a.T_COL_1);
            if (rowNum == 3) return tableData.Where(a => a.TableIndex == TableIndex.T_Row_Type_3).OrderByDescending(a => a.T_COL_1);
            return tableData.AsEnumerable();
        }
    }
}