using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace FaultReasoning
{
    public partial class Form1 : Form
    {
        System.Data.DataTable dt_symptom = new System.Data.DataTable();    // 故障征兆表 
        System.Data.DataTable dt_fault = new System.Data.DataTable();    // 故障表 
        _Workbook book;//电子表格
        int[,] matrix;//关联度矩阵
        public Form1()
        {
            InitializeComponent();
            try
            {
                //启动Excel应用程序
                Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
                //    _Workbook book = xls.Workbooks.Add(Missing.Value); //创建一张表，一张表可以包含多个sheet

                //如果表已经存在，可以用下面的命令打开
                book = xls.Workbooks.Open(System.Environment.CurrentDirectory + @"\content.xlsx");

                _Worksheet sheet;//定义sheet变量
                xls.Visible = false;//设置Excel后台运行
                xls.DisplayAlerts = false;//设置不显示确认修改提示
                sheet = (_Worksheet)book.Worksheets.get_Item(1);//获得第i个sheet，准备写入

                dt_symptom.Columns.Add(new DataColumn("isChecked", typeof(bool)));
                dt_symptom.Columns.Add(new DataColumn("序号", typeof(int)));
                dt_symptom.Columns.Add(new DataColumn("故障征兆", typeof(string)));
                dt_fault.Columns.Add(new DataColumn("序号", typeof(int)));
                dt_fault.Columns.Add(new DataColumn("故障", typeof(string)));
                //填充故障征兆表
                for (int index = 1; index < 300; index++)
                {
                    Range temp_range = sheet.Cells[index + 1, 1];
                    if (temp_range.Value == null)
                        break;
                    else
                        dt_symptom.Rows.Add(false, index, temp_range.Value);
                }
                //填充故障表
                for (int index = 1; index < 300; index++)
                {
                    Range temp_range = sheet.Cells[1, index + 1];
                    if (temp_range.Value == null)
                        break;
                    else
                        dt_fault.Rows.Add(index, temp_range.Value);
                }

                //填充规则矩阵
                matrix = new int[dt_symptom.Rows.Count, dt_fault.Rows.Count];
                for (int i = 0; i < dt_symptom.Rows.Count; i++)
                    for (int j = 0; j < dt_fault.Rows.Count; j++)
                    {
                        Range temp_range = sheet.Cells[i + 2, j + 2];
                        matrix[i, j] = Convert.ToInt32(temp_range.Value);
                    }
                dataGrid.DataSource = dt_symptom;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void OnValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void OnSubmit(object sender, EventArgs e)//提交代码处理
        {
            try
            {
                System.Data.DataTable temp_dt = dt_fault.Copy();
                for (int index = 0; index < dt_fault.Rows.Count; index++)
                {
                    string str_fault = dt_fault.Rows[index]["故障"] as string;
                    for (int i = 0; i < dt_symptom.Rows.Count; i++)
                    {
                        if (matrix[i, index] == 1)//规则1必要性检测
                        {
                            if ((bool)dt_symptom.Rows[i]["isChecked"] == false)
                            {
                                DataRow[] rows = temp_dt.Select("故障='" + str_fault + "'");
                                if (rows.Count() == 1)
                                {
                                    temp_dt.Rows.Remove(rows[0]);
                                    continue;
                                }
                                else if (rows.Count() > 1)
                                {
                                    throw new Exception("故障名称重复");
                                }
                            }
                        }
                        if (matrix[i, index] == -1)//规则3充分性检测
                        {
                            if ((bool)dt_symptom.Rows[i]["isChecked"] == true)
                            {
                                DataRow[] rows = temp_dt.Select("故障='" + str_fault + "'");
                                if (rows.Count() == 1)
                                {
                                    temp_dt.Rows.Remove(rows[0]);
                                    continue;
                                }
                                else if (rows.Count() > 1)
                                {
                                    throw new Exception("故障名称重复");
                                }
                            }
                        }
                    }
                }
                dataResult.DataSource = temp_dt;
                // temp_dt.Rows.
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void OnFormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                book.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
