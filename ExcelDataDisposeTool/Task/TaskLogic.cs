using System.Data;

//中转页
namespace ExcelDataDisposeTool.Task
{
    public class TaskLogic
    {
        Generate generate = new Generate();
        ImportDt importDt = new ImportDt();

        #region 变量参数
        private int _taskid;
        private string _fileAddress;        //文件地址('自定义批量导出'-导入EXCEL 及 导出地址收集使用)
        private DataTable _excelDt;         //获取前端的EXCEL导入DT

        private bool _resultmark;         //返回是否成功标记
        private DataTable _resultImportDt;  //返回导入EXCEL信息

        #endregion

        #region Set
        /// <summary>
        /// 中转ID
        /// </summary>
        public int TaskId { set { _taskid = value; } }
        /// <summary>
        /// 接收文件地址信息
        /// </summary>
        public string FileAddress { set { _fileAddress = value; } }

        /// <summary>
        /// 获取前端的EXCEL导入DT
        /// </summary>
        public DataTable ExcelDt { set { _excelDt = value; } }

        #endregion

        #region Get
        /// <summary>
        ///  返回是否成功标记
        /// </summary>
        public bool ResultMark => _resultmark;

        /// <summary>
        /// 返回导入EXCEL结果
        /// </summary>
        public DataTable ResultImportDt => _resultImportDt;
        #endregion


        public void StartTask()
        {
            switch (_taskid)
            {
                //导入
                case 0:
                    ImportExcelRecord(_fileAddress);
                    break;
                //运算
                case 1:
                    GenerateRecord(_excelDt, _fileAddress);
                    break;
            }
        }

        /// <summary>
        /// 导入EXCEL
        /// </summary>
        /// <param name="fileadd"></param>
        private void ImportExcelRecord(string fileadd)
        {
            _resultImportDt = importDt.OpenExcelImporttoDt(fileadd).Copy();
        }

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="excelDt">excel导入DT</param>
        /// <param name="fileAddress">导出设置地址</param>
        private void GenerateRecord(DataTable excelDt,string fileAddress)
        {
            _resultmark = generate.GenerateRecord(excelDt, fileAddress);
        }
    }
}
