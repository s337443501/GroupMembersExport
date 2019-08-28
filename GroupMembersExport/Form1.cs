using Newtonsoft.Json.Linq;
using SufeiUtil;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GroupMembersExport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FrmLoginQQ frmLoginQQ = new FrmLoginQQ(this);
            frmLoginQQ.ShowDialog();
        }

        public string bkn = "", cookies = "";
        private void button2_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                if (comboBox1.SelectedIndex == 0)
                {
                    CommonHelper.ExportExcel(dataGridView1);
                }
                else if (comboBox1.SelectedIndex == 1)
                {
                    ExportTxt(dataGridView1, textBox1.Text.Trim());
                }
            }
            else if (radioButton2.Checked)
            {
                if (comboBox1.SelectedIndex == 0)
                {
                    CommonHelper.ExportExcel(dataGridView2);
                }
                else if (comboBox1.SelectedIndex == 1)
                {
                    ExportTxt(dataGridView2, textBox1.Text.Trim());
                }
                else if (comboBox1.SelectedIndex == 2)
                {
                    try
                    {
                        if (dataGridView2.RowCount <= 0) return;
                        SaveFileDialog saveDialog = new SaveFileDialog();
                        saveDialog.DefaultExt = "txt";
                        saveDialog.Filter = "文本文档|*.txt";
                        saveDialog.FileName = "userqq";
                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            string filePath = saveDialog.FileName;
                            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                            {
                                StringBuilder sb = new StringBuilder();
                                foreach (DataGridViewRow dr in dataGridView2.Rows)
                                {
                                    sb.AppendLine(dr.Cells[1].Value.ToString());
                                }
                                sw.Write(sb.ToString());
                                MessageBox.Show("导出完成");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        public void ExportTxt(DataGridView dgv, string delimiter)
        {
            try
            {
                if (dgv.RowCount <= 0) return;
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.DefaultExt = "txt";
                saveDialog.Filter = "文本文档|*.txt";
                saveDialog.FileName = "";
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveDialog.FileName;
                    using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (DataGridViewRow dr in dgv.Rows)
                        {
                            string temp = string.Empty;
                            for (int i = 0; i < dgv.ColumnCount; i++)
                            {
                                temp += dr.Cells[i].Value + delimiter;
                            }
                            temp = temp.Substring(0, temp.Length - delimiter.Length);
                            sb.AppendLine(temp);
                        }
                        sw.Write(sb.ToString());
                        MessageBox.Show("导出完成");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var mess = "操作说明：\r\n1.先登录QQ再进行后续操作\r\n2.双击QQ群列表可以获取QQ群成员数据\r\n3.点击表格标题可以自定义排序";
            MessageBox.Show(mess);
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.RowCount > 0)
                {
                    int count = dataGridView1.SelectedRows.Count;
                    string groupid = "";
                    for (int i = 0; i < count; i++)
                    {
                        //string id = dataGridView1.SelectedRows[i].Cells["id"].Value.ToString();
                        //string groupname = dataGridView1.SelectedRows[i].Cells["groupname"].Value.ToString();
                        groupid += dataGridView1.SelectedRows[i].Cells["groupid"].Value.ToString() + "|";
                    }
                    Task.Factory.StartNew(() => ShowDT(groupid));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.RowCount > 0)
                {
                    //string id = dataGridView1.SelectedRows[0].Cells["id"].Value.ToString();
                    //string groupname = dataGridView1.SelectedRows[0].Cells["groupname"].Value.ToString();
                    string groupid = dataGridView1.SelectedRows[0].Cells["groupid"].Value.ToString();
                    Task.Factory.StartNew(() => ShowDT(groupid + "|"));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private delegate void UpdateDataGridView(DataTable dt);
        DataTable dataTable;
        private void UpdateGV(DataTable dt)
        {
            if (dataGridView1.InvokeRequired)
            {
                this.BeginInvoke(new UpdateDataGridView(UpdateGV), new object[] { dt });
            }
            else
            {
                dataGridView2.DataSource = dt;
                dataGridView2.Refresh();
            }
        }

        public void ShowDT(string groupid)
        {
            dataTable = new DataTable();
            dataTable.Columns.Add("id", typeof(int));
            dataTable.Columns.Add("uin", typeof(string)); //QQ号码
            dataTable.Columns.Add("nk", typeof(string)); //QQ昵称
            dataTable.Columns.Add("ll_lp", typeof(string)); //等级_积分
            dataTable.Columns.Add("jt", typeof(string)); //入群时间
            dataTable.Columns.Add("lst", typeof(string)); //最后发言时间
            string[] groupidArray = groupid.Split('|');
            for (int i = 0; i < groupidArray.Length - 1; i++)
            {
                SetDataTable(groupidArray[i]);
            }

            UpdateGV(dataTable);
        }

        private void SetDataTable(string groupid)
        {
            try
            {
                Dictionary<string, string> levelName = new Dictionary<string, string>();
                var url = "https://qinfo.clt.qq.com/cgi-bin/qun_info/get_members_info_v1?friends=1&name=1&src=qinfo_v3&gc=" + groupid + "&bkn=" + bkn;
                var html = CommonHelper.GetHtml(url, cookies);
                var jobj = JObject.Parse(html);
                var jo = JObject.Parse(jobj["levelname"].ToString());
                foreach (JProperty jProperty in jo.Properties())
                {
                    //Console.WriteLine(jProperty.Name + "：" + jProperty.Value);
                    levelName.Add(jProperty.Name.Replace("lvln", ""), jProperty.Value.ToString());
                }

                var jo1 = JObject.Parse(jobj["members"].ToString());
                int uid = 0;
                foreach (JProperty jProperty in jo1.Properties())
                {
                    string uin = jProperty.Name;
                    var j = JObject.Parse(jProperty.Value.ToString());
                    string nk = j["nk"].ToString();
                    string ll_lp = "";
                    string ll = j["ll"].ToString();
                    string lp = j["lp"].ToString();
                    string jt = j["jt"].ToString();
                    string lst = j["lst"].ToString();
                    nk = System.Web.HttpUtility.HtmlDecode(nk);
                    if (!string.IsNullOrEmpty(jt)) { jt = CommonHelper.NumToTime(jt); }
                    if (!string.IsNullOrEmpty(lst)) { lst = CommonHelper.NumToTime(lst); }
                    if (!string.IsNullOrEmpty(ll))
                    {
                        foreach (var lev in levelName)
                        {
                            if (ll == lev.Key)
                            {
                                ll = lev.Value;
                            }
                        }
                    }
                    ll_lp = ll + "(" + lp + ")";
                    uid++;
                    dataTable.Rows.Add(uid, uin, nk, ll_lp, jt, lst);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
