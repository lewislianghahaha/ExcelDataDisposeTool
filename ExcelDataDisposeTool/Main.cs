using System;
using System.Data;
using System.Threading;
using System.Windows.Forms;
using ExcelDataDisposeTool.Task;

namespace ExcelDataDisposeTool
{
    public partial class Main : Form
    {
        Load load=new Load();
        TaskLogic taskLogic=new TaskLogic();

        public Main()
        {
            InitializeComponent();
            OnRegisterEvents();
        }

        private void OnRegisterEvents()
        {
            tmclose.Click += Tmclose_Click;
            tmimport.Click += Tmimport_Click;
            btnsetadd.Click += Btnsetadd_Click;
        }

        /// <summary>
        /// 设置导出地址
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btnsetadd_Click(object sender, EventArgs e)
        {
            try
            {
                //todo:若txtadd不为空,即先清空此控件
                txtadd.Text = "";
                //todo:设置输出地址
                var folder = new FolderBrowserDialog();
                folder.Description = $"请选择导出文件夹";
                if (folder.ShowDialog() != DialogResult.OK) return;
                txtadd.Text = folder.SelectedPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, $"错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导入EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Tmimport_Click(object sender, EventArgs e)
        {
            try
            {
                //todo:检测到若导出地址为空,不能继续
                if(txtadd.Text == "") throw new Exception("导出地址不能空,请选择后继续");

                //设置导入地址
                var openFileDialog = new OpenFileDialog { Filter = $"Xlsx文件|*.xlsx" };
                if (openFileDialog.ShowDialog() != DialogResult.OK) return;
                var fileAdd = openFileDialog.FileName;

                //所需的值赋到Task类内
                taskLogic.TaskId = 0;
                taskLogic.FileAddress = fileAdd;


                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();

                var importdt = taskLogic.ResultImportDt.Copy();

                if (importdt.Rows.Count == 0) throw new Exception("不能成功导入EXCEL内容,请检查模板是否正确.");
                else
                {
                    //导入EXCEL记录完成后,进入运算
                    if (!GenerateExcelDt(importdt))
                    {
                        MessageBox.Show($"运算出现异常,请联系管理员", $"信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        MessageBox.Show($"运算及导出成功,请到'{txtadd.Text}'地址进行查阅", $"信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                //无论成功与否,结束后都会文本框清空
                txtadd.Text = "";
                tmclose.Enabled = true;
                tmimport.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, $"错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="importdt">导入DT</param>
        /// <returns></returns>
        private bool GenerateExcelDt(DataTable importdt)
        {
            var result = true;

            try
            {
                var message = $"准备执行,\n请注意:" +
                              $"\n1.执行成功的结果会下载至'{txtadd.Text}'指定文件夹内," +
                              "\n2.执行过程中不要关闭软件,不然会导致运算失败" + "\n是否继续执行?";

                if (MessageBox.Show(message, $"提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    tmclose.Enabled = false;
                    tmimport.Enabled = false;

                    //所需的值赋到Task类内
                    taskLogic.TaskId = 1;
                    taskLogic.ExcelDt = importdt;
                    taskLogic.FileAddress = txtadd.Text;

                    //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                    new Thread(Start).Start();
                    load.StartPosition = FormStartPosition.CenterScreen;
                    load.ShowDialog();

                    result = taskLogic.ResultMark;

                }
            }
            catch (Exception)
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        /// 关闭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Tmclose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        ///子线程使用(重:用于监视功能调用情况,当完成时进行关闭LoadForm)
        /// </summary>
        private void Start()
        {
            taskLogic.StartTask();

            //当完成后将Form2子窗体关闭
            this.Invoke((ThreadStart)(() =>
            {
                load.Close();
            }));
        }

    }
}
