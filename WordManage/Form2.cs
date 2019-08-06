using System;
using System.Collections;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace WordManage
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        #region 显示要操作的文件
        /// <summary>
        /// 显示要操作的文件
        /// </summary>
        /// <param name="strPath">路径</param>
        void ShowFile(string strPath)
        {
            listView1.Items.Clear();//清空ListView列表项
            string[] files = Directory.GetFiles(strPath);//从路径中获取所有文件
            foreach (string file in files)//遍历所有文件
            {
                FileInfo finfo = new FileInfo(file);//创建文件对象
                listView1.Items.Add(finfo.Name);//获取文件名并显示
                listView1.Items[listView1.Items.Count - 1].SubItems.Add(finfo.DirectoryName);//获取路径并显示
            }
        }
        #endregion

        //选择要操作的Word所在路径
        private void Button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dir = new FolderBrowserDialog();//创建浏览文件对话框
            if (dir.ShowDialog() == DialogResult.OK)//判断是否选择了路径
            {
                textBox1.Text = dir.SelectedPath;//显示选择的路径
                ShowFile(textBox1.Text);//调用方法显示路径下的所有文件
            }
        }

        //设置要重命名的模板txt文件
        private void Button2_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();//清空文本框
            OpenFileDialog openfile = new OpenFileDialog();//创建打开文件对话框
            openfile.Filter = "txt模板文件|*.txt";//设置只能打开txt文件
            if (openfile.ShowDialog() == DialogResult.OK)//判断是否选择了文件
            {
                textBox2.Text = openfile.FileName;//显示打开的文件名
                StreamReader SReader = new StreamReader(textBox2.Text, Encoding.UTF8);//以UTF8编码方式读取文件
                richTextBox1.Text = SReader.ReadToEnd();//读取所有文件并显示在文本框中
            }
        }

        //执行Word文档整理工具（按照选择的txt文件对Word文件名进行重命名）
        private void Button3_Click(object sender, EventArgs e)
        {
            string[] files = Directory.GetFiles(textBox1.Text);//获取所有需要重命名的Word文件
            int i = 0;//定义一个标识，用来作为遍历时的文件索引
            ArrayList list = new ArrayList();//创建List集合，用来存储新文件名
            foreach (string item in richTextBox1.Lines)//遍历文本框中的所有行
                list.Add(item);//将所有行添加到集合中进行存储
            foreach(string str in list)//遍历存储新文件名的集合
            {
                try
                {
                    //依次对遍历到的文件进行重命名
                    File.Copy(files[i], Path.GetDirectoryName(files[i]).TrimEnd(new char[] { '\\' }) + "\\" + str + Path.GetExtension(files[i]), true);
                    File.Delete(files[i]);//删除原有文件
                    i++;//遍历文件的索引加1
                }
                catch//捕获异常（防止有模板中新文件名个数与要操作的Word文档个数不同的情况）
                {
                    MessageBox.Show("请重新选择文件路径！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            ShowFile(textBox1.Text);//显示重命名完成后的所有文件
            //成功提示
            MessageBox.Show("文件重命名整理完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //打开重命名后的路径进行查看
        private void Button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(textBox1.Text);//打开重命名后的路径进行查看
        }
    }
}