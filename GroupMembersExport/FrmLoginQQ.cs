using System;
using System.Data;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace GroupMembersExport
{
    public partial class FrmLoginQQ : Form
    {
        Form1 f1 = new Form1();
        public FrmLoginQQ(Form1 form1)
        {
            InitializeComponent();
            f1 = form1;
        }

        private void FrmLoginQQ_Load(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://xui.ptlogin2.qq.com/cgi-bin/xlogin?pt_disable_pwd=1&appid=715030901&daid=73&hide_close_icon=1&pt_no_auth=1&s_url=https://qun.qq.com/member.html");
            webBrowser1.ProgressChanged += WebBrowser1_ProgressChanged;
        }

        private void WebBrowser1_ProgressChanged(object sender, WebBrowserProgressChangedEventArgs e)
        {
            try
            {
                if (webBrowser1.Url != null)
                {
                    if (webBrowser1.Url.ToString().StartsWith("https://qun.qq.com/member.html"))
                    {
                        webBrowser1.ProgressChanged -= WebBrowser1_ProgressChanged;
                        string cookies = webBrowser1.Document.Cookie;
                        string Cookie_skey = CommonHelper.GetCookieValue(cookies, "skey");
                        string bkn = CommonHelper.Getbkn(Cookie_skey);
                        string url = "https://qun.qq.com/cgi-bin/qunwelcome/myinfo?callback=?&bkn=" + bkn;
                        var html = CommonHelper.GetHtml(url, cookies);
                        var jobj = JObject.Parse(html);
                        string uin = jobj["data"]["uin"].ToString();
                        string nickName = jobj["data"]["nickName"].ToString();
                        f1.label1.Text = "QQ号码：" + uin;
                        f1.label2.Text = "昵称：" + nickName;
                        f1.bkn = bkn;
                        f1.cookies = cookies;

                        url = "https://qun.qq.com/cgi-bin/qun_mgr/get_group_list?bkn=" + bkn;
                        html = CommonHelper.GetHtml(url, cookies);
                        jobj = JObject.Parse(html);

                        DataTable dataTable = new DataTable();
                        dataTable.Columns.Add("id", typeof(int));
                        dataTable.Columns.Add("groupname", typeof(string));
                        dataTable.Columns.Add("groupid", typeof(string));
                        var jarr = JArray.Parse(jobj["join"].ToString());
                        int count = 1;
                        for (var i = 0; i < jarr.Count; i++)
                        {
                            var j = JObject.Parse(jarr[i].ToString());
                            string groupname = j["gn"].ToString();
                            string groupid = j["gc"].ToString();
                            dataTable.Rows.Add(count, groupname, groupid);
                            count++;
                        }
                        jarr = JArray.Parse(jobj["manage"].ToString());
                        for (var i = 0; i < jarr.Count; i++)
                        {
                            var j = JObject.Parse(jarr[i].ToString());
                            string groupname = j["gn"].ToString();
                            string groupid = j["gc"].ToString();
                            dataTable.Rows.Add(count, groupname, groupid);
                            count++;
                        }
                        jarr = JArray.Parse(jobj["create"].ToString());
                        for (var i = 0; i < jarr.Count; i++)
                        {
                            var j = JObject.Parse(jarr[i].ToString());
                            string groupname = j["gn"].ToString();
                            string groupid = j["gc"].ToString();
                            dataTable.Rows.Add(count, groupname, groupid);
                            count++;
                        }
                        f1.dataGridView1.DataSource = dataTable;

                        webBrowser1.Dispose(); this.Close();
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
