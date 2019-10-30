using System;
using System.Collections;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace WordManage
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        #region 显示要操作的文件
        /// <summary>
        /// 显示要操作的文件
        /// </summary>
        /// <param name="strPath">路径</param>
        void ShowFile(string strPath)
        {
            listView1.Items.Clear();//清空ListView列表项
            if (Directory.Exists(strPath))//判断路径是否存在
            {
                string[] files = Directory.GetFiles(strPath);//从路径中获取所有文件
                foreach (string file in files)//遍历所有文件
                {
                    FileInfo finfo = new FileInfo(file);//创建文件对象
                    if (finfo.Extension == ".doc" || finfo.Extension == ".docx")//判断是否为Word文件
                    {
                        listView1.Items.Add(finfo.Name);//获取文件名并显示
                        listView1.Items[listView1.Items.Count - 1].SubItems.Add(finfo.DirectoryName);//获取路径并显示
                    }
                }
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
                richTextBox1.Text = SReader.ReadToEnd();//读取所有文件内容并显示在文本框中
                SReader.Close();
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
            if (listView1.Items.Count == list.Count) //判断要重命名的文件与模板中的新文件名行数一致
            {
                button3.Enabled = false;
                foreach (string str in list)//遍历存储新文件名的集合
                {
                    //依次对遍历到的文件进行重命名
                    File.Copy(files[i], Path.GetDirectoryName(files[i]).TrimEnd(new char[] { '\\' }) + "\\" + str + Path.GetExtension(files[i]), true);
                    File.Delete(files[i]);//删除原有文件
                    i++;//遍历文件的索引加1
                }
                ShowFile(textBox1.Text);//显示重命名完成后的所有文件
                button3.Enabled = true;
                //成功提示
                MessageBox.Show("文件重命名整理完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("文件个数与模板文件中的行数不一致，请重新确认！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        //打开重命名后的路径进行查看
        private void Button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(textBox1.Text);//打开重命名后的路径进行查看
        }
        ArrayList listDoc=new ArrayList(), listTxt = new ArrayList();
        private void Button5_Click(object sender, EventArgs e)
        {
            listDoc.Clear();
            FolderBrowserDialog dir = new FolderBrowserDialog();//创建浏览文件对话框
            if (dir.ShowDialog() == DialogResult.OK)//判断是否选择了路径
            {
                textBox3.Text = dir.SelectedPath;//显示选择的路径
                //存储所有文件夹路径
                foreach (string str in Directory.GetDirectories(textBox3.Text))
                    listDoc.Add(new DirectoryInfo(str).Name);
            }
            
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            listTxt.Clear();
            FolderBrowserDialog dir = new FolderBrowserDialog();//创建浏览文件对话框
            if (dir.ShowDialog() == DialogResult.OK)//判断是否选择了路径
            {
                textBox4.Text = dir.SelectedPath;//显示选择的路径
                //存储所有重命名模板文件
                foreach (string str in Directory.GetFiles(textBox4.Text, "*.txt"))
                    listTxt.Add(Path.GetFileNameWithoutExtension(str));
            }
        }

        private void Button8_Click(object sender, EventArgs e)
        {

            if (listDoc != null && listTxt != null)
            {
                button8.Enabled = false;
                foreach (string str in listDoc)
                {
                    if (listTxt.Contains(str))
                    {
                        int i = 0;//定义一个标识，用来作为遍历时的文件索引
                        StreamReader SReader = new StreamReader(textBox4.Text.TrimEnd('\\') + "\\" + str + ".txt", Encoding.UTF8);//以UTF8编码方式读取文件
                        string[] files = Directory.GetFiles(textBox3.Text.TrimEnd('\\') + "\\" + str, "*.doc");//获取所有需要重命名的Word文件
                        Array.Sort(files);
                        string strLine = string.Empty; //记录每次读到的行
                        while ((strLine = SReader.ReadLine()) != null) //循环按行读取内容
                        {
                            //依次对遍历到的文件进行重命名
                            File.Copy(files[i], Path.GetDirectoryName(files[i]).TrimEnd(new char[] { '\\' }) + "\\" + strLine + Path.GetExtension(files[i]), true);
                            File.Delete(files[i]);//删除原有文件
                            i++;
                        }
                        SReader.Close(); //关闭读取器
                    }
                }
                button8.Enabled = true;
                //成功提示
                MessageBox.Show("文件重命名整理完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter) 
            {
                ShowFile(textBox1.Text);//调用方法显示路径下的所有文件
            }
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(textBox3.Text);//打开重命名后的路径进行查看
        }
    }
}