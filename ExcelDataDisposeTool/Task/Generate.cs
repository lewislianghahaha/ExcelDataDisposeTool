using System;
using System.Data;
using System.Windows.Forms;
using ExcelDataDisposeTool.DB;
using Stimulsoft.Report;

//运算
namespace ExcelDataDisposeTool.Task
{
    public class Generate
    {
        TempDtList tempDtList = new TempDtList();

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="excelDt"></param>
        /// <param name="fileAddress"></param>
        /// <returns></returns>
        public bool GenerateRecord(DataTable excelDt, string fileAddress)
        {
            var result = true;

            try
            {
                //todo:将excelDt进行数据处理,将含有"汇总" 及 “总计”的行数去除,只保留明细行
                var tempdt = GetNewDt(excelDt.Clone(),excelDt).Copy();
                //todo:根据筛选出来的值-获取不重复的'交货单号'记录集
                var distdt = GetDistDt(tempdt).Copy();
                //todo:根据distdt 以及 tempdt进行循环生成PDF输出
                result = ExportDtToPdf(distdt,fileAddress,tempdt);
            }
            catch (Exception ex)
            {
                var a = ex.Message;
                result = false;
            }

            return result;
        }

        /// <summary>
        /// 根据筛选出来的值-获取不重复的'交货单号'记录集
        /// </summary>
        /// <param name="tempdt"></param>
        /// <returns></returns>
        private DataTable GetDistDt(DataTable tempdt)
        {
            var temp = string.Empty;
            var resuldt = tempDtList.DistData();

            foreach (DataRow rows in tempdt.Rows)
            {
                if (string.IsNullOrEmpty(temp))
                {
                    var newrow = resuldt.NewRow();
                    newrow[0] = Convert.ToString(rows[0]);
                    temp = Convert.ToString(rows[0]);
                    resuldt.Rows.Add(newrow);
                }
                else if (temp !=Convert.ToString(rows[0]))
                {
                    var newrow = resuldt.NewRow();
                    newrow[0] = Convert.ToString(rows[0]);
                    temp = Convert.ToString(rows[0]);
                    resuldt.Rows.Add(newrow);
                }
            }

            return resuldt;
        }

        /// <summary>
        /// 根据从EXCELDT记录集进行整理，去除 将含有"汇总" 及 “总计”的行数去除,只保留明细行
        /// </summary>
        /// <returns></returns>
        private DataTable GetNewDt(DataTable tempdt,DataTable exceldt)
        {
            var dtlrows = exceldt.Select("FORDERNO not like '%汇总%' and FORDERNO <>'总计'");

            if (dtlrows.Length > 0)
            {
                for (var i = 0; i < dtlrows.Length; i++)
                {
                    var newrow = tempdt.NewRow();
                    newrow[0] = dtlrows[i][0];   //交货单号
                    newrow[1] = dtlrows[i][1];   //交货单行项目
                    newrow[2] = dtlrows[i][2];   //物料号
                    newrow[3] = dtlrows[i][3];   //物料描述
                    newrow[4] = dtlrows[i][4];   //交货数量
                    newrow[5] = dtlrows[i][5];   //物料编码
                    newrow[6] = dtlrows[i][6];   //单价
                    newrow[7] = dtlrows[i][7];   //价税合计
                    newrow[8] = dtlrows[i][8];   //客户名称
                    newrow[9] = dtlrows[i][9];   //收件人姓名
                    newrow[10] = dtlrows[i][10]; //收件人电话
                    newrow[11] = dtlrows[i][11]; //收件地址
                    tempdt.Rows.Add(newrow);
                }
            }

            return tempdt;    
        }


        /// <summary>
        /// 将计算的DT输出至指定位置的PDF-循环输出
        /// </summary>
        /// <param name="distdt">‘交货单号’唯一DT</param>
        /// <param name="exportaddress">输出地址</param>
        /// <param name="resultdt">结果集</param>
        /// <returns></returns>
        private bool ExportDtToPdf(DataTable distdt,string exportaddress, DataTable resultdt)
        {
            var result = true;
            var stiReport = new StiReport();
            var dt = DateTime.Now.ToString("yyyy-MM-dd");

            try
            {
                var filepath = Application.StartupPath + "/Report/" + "ExcelRecordShow.mrt";

                //todo:通过循环distdt,将整合各记录集并分文件输出
                foreach (DataRow rows in distdt.Rows)
                {
                    var finalresultdt = MakeResultdt(Convert.ToString(rows[0]),resultdt).Copy();
                     
                    var pdfFileAddress = exportaddress + "\\"+ $"交货单号_{Convert.ToString(rows[0])}_输出记录_" + $"{dt}" + ".pdf";

                    stiReport.Load(filepath);                                      //读取STI模板
                    stiReport.RegData("ExcelRecordShow", finalresultdt);           //填充数据至STI模板内
                    stiReport.Render(false);                                       //重点-没有这个生成的文件会提示“文件已损坏”
                    stiReport.ExportDocument(StiExportFormat.Pdf, pdfFileAddress); //生成指定格式文件   
                }
            }
            catch (Exception ex)
            {
                var a = ex.Message;
                result = false;
            }

            return result;
        }

        /// <summary>
        /// 根据不同的‘交货单号’整合记录集
        /// </summary>
        /// <param name="forderno"></param>
        /// <param name="resultdt"></param>
        /// <returns></returns>
        private DataTable MakeResultdt(string forderno,DataTable resultdt)
        {
            var tempdt = resultdt.Clone();

            var dtlrows = resultdt.Select("FORDERNO='" + forderno + "'");

            for (var i = 0; i < dtlrows.Length; i++)
            {
                var newrow = tempdt.NewRow();
                newrow[0] = dtlrows[i][0];   //交货单号
                newrow[1] = dtlrows[i][1];   //交货单行项目
                newrow[2] = dtlrows[i][2];   //物料号
                newrow[3] = dtlrows[i][3];   //物料描述
                newrow[4] = dtlrows[i][4];   //交货数量
                newrow[5] = dtlrows[i][5];   //物料编码
                newrow[6] = dtlrows[i][6];   //单价
                newrow[7] = dtlrows[i][7];   //价税合计
                newrow[8] = dtlrows[i][8];   //客户名称
                newrow[9] = dtlrows[i][9];   //收件人姓名
                newrow[10] = dtlrows[i][10]; //收件人电话
                newrow[11] = dtlrows[i][11]; //收件地址
                tempdt.Rows.Add(newrow);
            }

            return tempdt;
        }


    }
}
