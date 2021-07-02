using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;

namespace 爬爬爬
{
    //界面层
    public partial class Form1 : Form
    {
        //建立一个委托，去刷新主线程UI
        delegate void AsynUpdateUI(int step,string str,int max);
        delegate void AsynComplete();

        DataWrite dataWrite = new DataWrite();//实例化一个写入数据的类

        public Form1()
        {
            InitializeComponent();
            textBox2.Text = "5";
            input_box.Text = "fur";
            button2.Enabled = false;
            progressBar1.Visible = false;
            progressBar1.Value = 20;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (int.TryParse(textBox2.Text, out int page) == false || page <=0)
            {
                MessageBox.Show("请输入合适的页数");
                return;
            }
            textBox1.Text = "";

            //初始化进度条
            progressBar1.Visible = true;
            progressBar1.Maximum = page;//初始化进度条任务数量
            progressBar1.Value = 0;

            string serchStr = input_box.Text;

            pram pr = new pram();
            pr.page = page;
            pr.searchStr = serchStr;

            dataWrite = new DataWrite();
            dataWrite.UpdateUIDelegate += UpdataUIStatus;//绑定更新任务状态的委托
            dataWrite.TaskCallBack += Accomplish;//绑定完成任务要调用的委托
            this.button1.Enabled = false;
            Thread thread = new Thread(new ParameterizedThreadStart(dataWrite.start));
            thread.IsBackground = true;
            thread.Start(pr);
        }
        //存到excel
        
        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            dataWrite.saveToExcel(path.SelectedPath);
            this.button2.Enabled = false;
        }
        //更新UI
        private void UpdataUIStatus(int step,string str,int max)
        {
            if (InvokeRequired)
            {
                this.Invoke(new AsynUpdateUI(delegate (int s,string st,int mx)
                {
                    if (mx == 0)
                    {
                        this.progressBar1.Visible = false;
                    }
                    else
                    {
                        this.progressBar1.Visible = true;
                        this.progressBar1.Value = s;
                        this.progressBar1.Maximum = mx;
                        this.progressBar1.Text = s.ToString() + "/" + mx.ToString();
                    }
                    this.textBox1.Text += st;
                }), new object[] { step, str,max});
            }
            else
            {
                if (max == 0)
                {
                    this.progressBar1.Visible = false;
                }
                else
                {
                    this.progressBar1.Visible = true;
                    this.progressBar1.Value = step;
                    this.progressBar1.Maximum = max;
                    this.progressBar1.Text = step.ToString() + "/" + max.ToString();
                }
                this.textBox1.Text += str;
            }
        }

        //完成任务时需要调用
        private void Accomplish()
        {
            if (InvokeRequired)
            {
                this.Invoke(new AsynComplete(delegate ()
                {
                    MessageBox.Show("爬取完成，可以进行保存！");
                    //重置按钮
                    this.button1.Enabled = true;
                    this.button2.Enabled = true;
                }), new object[] { });
            }
            else
            {
                this.button1.Enabled = true;
                this.button2.Enabled = true;
                MessageBox.Show("爬取完成，可以进行保存！");
            }
          
        }
   
    }
    //逻辑层
    public class DataWrite
    {
        public delegate void UpdateUI(int step,string str,int max);//声明一个更新主线程的委托
        public UpdateUI UpdateUIDelegate;

        public delegate void AccomplishTask();//声明一个在完成任务时通知主线程的委托
        public AccomplishTask TaskCallBack;


        List<string> itemList = new List<string>();
        List<string> URLList = new List<string>();
        List<string> TitleList = new List<string>();
        Dictionary<string, infos> dic = new Dictionary<string, infos>();

        //异步的开始任务
        public void start(object obj)
        {
            nextThreadDoSth(obj);
            //任务完成时通知主线程作出相应的处理
            TaskCallBack();
        }
        //爬取与数据处理的主要函数
        private void nextThreadDoSth(object obj)
        {
            pram pr = obj as pram;
            if (pr == null)
            {
                MessageBox.Show("can not find object :'obj' on func 'nextThreadDoSth'");
                return;
            }

            string serchStr = pr.searchStr;
            int page = pr.page;

            dic.Clear();
            try
            {
                string baseUrl = "https://www.alibaba.com/products/searchStr.html?IndexArea=product_en&page=";
                serchStr.Replace(',', '_');
                serchStr.Replace(' ', '_');
                serchStr.Replace('，', '_');
                string searchUrl = baseUrl.Replace("searchStr", serchStr);
                HttpRequestUtil httpReq = new HttpRequestUtil();

                itemList = new List<string>();
                URLList = new List<string>();
                TitleList = new List<string>();
                //this.textBox1.Text += "已经开始爬取网页，请稍等········\r\n";
                UpdateUIDelegate(0, "已经开始爬取网页，请稍等········\r\n",0);
                List<string> list = new List<string>();
                for (int i = 1; i <= page; i++)
                {
                    string res = httpReq.GetPageHtml2(searchUrl + i);
                    //获取到主网页，需要从主页里解析出<div class=organic-list app-organic-search__list"
                    UpdateUIDelegate(i, "", page);
                    getItemList(res);
                }
                //this.textBox1.Text += "正在解析网页内容，请稍等········\r\n";
                UpdateUIDelegate(page, "正在解析网页内容，请稍等········\r\n", page);
                removeAdAndAddUrl();

                //遍历URL
                int index = 0;
                foreach (var URL in URLList)
                {
                    string URLRes = httpReq.GetPageHtml2("https://" + URL);
                    getTitle(URLRes);
                    UpdateUIDelegate(index, "", URLList.Count);
                    index++;
                }

                getAllInfos();
                //this.textBox1.Text += "==================================================\r\n";
                UpdateUIDelegate(0, "==================================================\r\n",0);
                //this.textBox1.Text += "下面是详细信息\r\n";
                UpdateUIDelegate(0, "下面是详细信息\r\n", 0);
                foreach (var var in dic)
                {
                    //this.textBox1.Text += "网址：" + var.Key + "\r\n";
                    UpdateUIDelegate(0, "网址：" + var.Key + "\r\n", 0);
                    //this.textBox1.Text += "名字：" + var.Value.name + "\r\n";
                    UpdateUIDelegate(0, "名字：" + var.Value.name + "\r\n", 0);
                    //this.textBox1.Text += "关键字：";
                    UpdateUIDelegate(0, "关键字：", 0);

                    foreach (var word in var.Value.words)
                    {
                        //this.textBox1.Text += word + "     ";
                        UpdateUIDelegate(0, word + "     ", 0);
                    }
                    //this.textBox1.Text += "\r\n";
                    UpdateUIDelegate(0, "\r\n", 0);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("发生了一些问题：" + exception);
            }
        }
        private void getItemList(string res)
        {
            int index = res.IndexOf("<div class=\"organic-list app-organic-search__list\"");
            int start = index;
            int end = start + 1;
            int all = res.Length;
            int flag = 0;
            while (end < all)
            {
                if (res[end] == '<' && res[end + 1] == 'd' && res[end + 2] == 'i' && res[end + 3] == 'v')
                {
                    flag += 1;
                }
                if (res[end] == '<' && res[end + 1] == '/' && res[end + 2] == 'd' && res[end + 3] == 'i' &&
                    res[end + 4] == 'v' && res[end + 5] == '>')
                {
                    flag -= 1;
                }

                if (flag < 0)
                    break;
                end++;
            }

            if (start < 0 || end < 0)
                return;
            string retStr = res.Substring(start, end - start + 6);
            //先取出所有的数据，然后根据当前数据去分割为一个一个的list
            int itemStart = start + 1;
            flag = -1;
            //找到下一个div的开头
            while (flag == -1)
            {
                if (res[itemStart] == '<' && res[itemStart + 1] == 'd' && res[itemStart + 2] == 'i' && res[itemStart + 3] == 'v')
                {
                    flag += 1;
                    break;
                }
                itemStart++;
            }
            int itemEnd = itemStart + 1;
            while (itemEnd < end)
            {
                if (res[itemEnd] == '<' && res[itemEnd + 1] == 'd' && res[itemEnd + 2] == 'i' && res[itemEnd + 3] == 'v')
                {
                    flag += 1;
                }
                if (res[itemEnd] == '<' && res[itemEnd + 1] == '/' && res[itemEnd + 2] == 'd' && res[itemEnd + 3] == 'i' &&
                    res[itemEnd + 4] == 'v' && res[itemEnd + 5] == '>')
                {
                    flag -= 1;
                }

                if (flag < 0)
                {
                    flag = -1;
                    itemList.Add(res.Substring(itemStart, itemEnd - itemStart + 6));
                    itemStart = itemEnd;
                    while (flag == -1)
                    {
                        if (res[itemStart] == '<' && res[itemStart + 1] == 'd' && res[itemStart + 2] == 'i' && res[itemStart + 3] == 'v')
                        {
                            flag += 1;
                            break;
                        }
                        itemStart++;
                    }
                    itemEnd = itemStart + 1;
                    continue;
                }
                else
                {
                    itemEnd++;
                }
            }
            // int num = Regex.Matches(retStr, "J-offer-wrapper").Count;
            return;
        }
        private void removeAdAndAddUrl()
        {
            List<int> removeList = new List<int>();
            for (int i = 0; i < itemList.Count; i++)
            {
                if (itemList[i].Contains("class=\"search-ad-icon\""))
                    removeList.Add(i);
            }

            for (int i = removeList.Count - 1; i >= 0; i--)
            {
                itemList.RemoveAt(removeList[i]);
            }
            //href="//www.alibaba
            foreach (var item in itemList)
            {
                int index = item.IndexOf("<a href=\"//www.alibaba");
                int end = index;
                while (item[end] != '\"' || item[end + 1] != ' ')
                {
                    end++;
                }
                URLList.Add(item.Substring(index + 11, end - index - 11));
            }
        }
        private void getTitle(string res)
        {
            int index = res.IndexOf("<title>");
            int end = res.IndexOf("</title>");
            TitleList.Add(res.Substring(index + 7, end - index - 22));
        }
        //拆分标题
        private void getAllInfos()
        {
            foreach (var var in TitleList)
            {
                int index = 0;
                int end = var.IndexOf("- Buy");
                string name = var.Substring(index, end - index);
                string wordStr = var.Substring(end + 5, var.Length - end - 5);
                List<string> words = wordStr.Split(',').ToList();
                if (dic.ContainsKey(URLList[TitleList.IndexOf(var)]))
                {
                    dic[URLList[TitleList.IndexOf(var)]].name = name;
                    dic[URLList[TitleList.IndexOf(var)]].words.AddRange(words);
                }
                else
                {
                    infos temp = new infos();
                    temp.words = words;
                    temp.name = name;

                    dic.Add(URLList[TitleList.IndexOf(var)], temp);
                }
            }
        }
        public void saveToExcel(string path)
        {
            path = path.ToLower();
            HSSFWorkbook workbook = new HSSFWorkbook();
            //创建工作表
            var sheet = workbook.CreateSheet("信息表");
            //创建标题行（重点）
            var row = sheet.CreateRow(0);
            //创建单元格
            var cellid = row.CreateCell(0);
            cellid.SetCellValue("网址");
            var cellname = row.CreateCell(1);
            cellname.SetCellValue("标题");
            var cellpwd = row.CreateCell(2);
            cellpwd.SetCellValue("关键字");
            int i = 1;
            foreach (var var in dic)
            {
                row = sheet.CreateRow(i);
                cellid = row.CreateCell(0);
                cellid.SetCellValue(var.Key);
                cellname = row.CreateCell(1);
                cellname.SetCellValue(var.Value.name);
                for (int line = 0; line < var.Value.words.Count; line++)
                {
                    var cell = row.CreateCell(line + 2);
                    cell.SetCellValue(var.Value.words[line]);
                }
                i++;
            }
            TimeSpan ts1 = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);
            string name = path + "\\商品信息" + ts1.TotalSeconds + ".xls";
            FileStream file = new FileStream(name, FileMode.OpenOrCreate, FileAccess.Write);
            workbook.Write(file);
            file.Dispose();
            MessageBox.Show("已保存于" + name);
        }
    }
    //数据结构层
    public class infos
    {
        public string name;
        public List<string> words;

        public infos()
        {
            name = "";
            words = new List<string>();
        }
    }
    public class pram
    {
        public string searchStr = "";
        public int page = 0;
    }

    

}
