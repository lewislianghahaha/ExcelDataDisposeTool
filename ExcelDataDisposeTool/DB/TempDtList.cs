using System;
using System.Data;

//临时表
namespace ExcelDataDisposeTool.DB
{
    public class TempDtList
    {
        /// <summary>
        /// 导出STI报表使用
        /// </summary>
        /// <returns></returns>
        public DataTable ExportData()
        {
            var dt = new DataTable();
            for (var i = 0; i < 12; i++)
            {
                var dc = new DataColumn();
                switch (i)
                {
                    //交货单号
                    case 0:
                        dc.ColumnName = "FORDERNO";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //交货单行项目
                    case 1:
                        dc.ColumnName = "FID";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //物料号
                    case 2:
                        dc.ColumnName = "FMATERICALNUM";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //物料描述
                    case 3:
                        dc.ColumnName = "FMATERIALDESC";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //交货数量
                    case 4:
                        dc.ColumnName = "FQTY";
                        dc.DataType = Type.GetType("System.Int32"); 
                        break;
                    //物料编码
                    case 5:
                        dc.ColumnName = "FMATERIALCODE";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //单价
                    case 6:
                        dc.ColumnName = "FPRICE";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //价税合计
                    case 7:
                        dc.ColumnName = "FTOTAL";
                        dc.DataType = Type.GetType("System.Double"); 
                        break;
                    //客户名称
                    case 8:
                        dc.ColumnName = "FCUSTOMNAME";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //收件人姓名
                    case 9:
                        dc.ColumnName = "FRECEIVENAME";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //收件人电话
                    case 10:
                        dc.ColumnName = "FRECEIVETEL";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //收件地址
                    case 11:
                        dc.ColumnName = "FRECEIVEADD";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 获取唯一的“交货单号”记录集
        /// </summary>
        /// <returns></returns>
        public DataTable DistData()
        {
            var dt = new DataTable();
            for (var i = 0; i < 1; i++)
            {
                var dc = new DataColumn();
                switch (i)
                {
                    //交货单号
                    case 0:
                        dc.ColumnName = "FORDERNO";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

    }
}
