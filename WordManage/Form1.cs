using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Threading;
using System.Diagnostics;

/*
 * Word文档拆分工具
 * 按书签拆分：手动在Word中把需要拆分的内容设置为书签
 * 按分页符拆分：适合没有表格的Word文档
 */
namespace WordManage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;//为了能够跨线程调用控件
        }

        int flag = 1;//标识按书签拆分还是分页符拆分，默认1，表示按分页符拆分
        ApplicationClass word = new ApplicationClass();//创建Word文档对象

        #region 读取Word文档
        /// <summary>
        /// 读取Word文档
        /// </summary>
        /// <param name="path">Word文档路径</param>
        /// <returns>Word文档对象</returns>
        private Document ReadDocument(string path)
        {
            Documents docs = null;//声明一个Word文档对象
            if(docs==null)//判断Word文档对象是否为空
                docs = word.Documents;//如果为空，则对其进行实例化
            Type docsType = docs.GetType();//获取Word文档对象的类型
            object objDocName = path;//定义一个变量，用来存储Word文档路径
            //打开Word文档对象
            Document doc = (Document)docsType.InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, docs, new Object[] { objDocName, true, true });
            return doc;//返回Word文档对象
        }
        #endregion

        #region 创建Word文档
        /// <summary>
        /// 创建Word文档
        /// </summary>
        /// <param name="normal">Word模板</param>
        /// <returns>Word文档对象</returns>
        private Document CreateDocument(object normal)
        {
            object missing = System.Reflection.Missing.Value;//作为方法传参中的缺省值
            //根据narmal指定的模板，使用Add方法添加Word文档
            Document newdoc = word.Documents.Add(ref normal, ref missing, ref missing, ref missing);
            return newdoc;//返回创建的Word文档对象
        }
        #endregion

        #region 获取书签的位置
        /// <summary>
        /// 获取书签的位置
        /// </summary>
        /// <param name="word">Word文档对象</param>
        /// <returns>存储Word文档中每个书签开始和结束位置的数组</returns>
        private int[,] GetPosition(Document doc)
        {
            int bmcount = doc.Bookmarks.Count;//获取书签的数量
            int[,] result = new int[bmcount, 2];//定义二维数组，用来存储书签的开始和结束位置
            int titleIndex = 0;//二维数组的索引（由于书签索引从1开始，而数组索引从0开始，所以定义该标识）
            for (int i = 1; i <= bmcount; i++)//遍历书签
            {
                object index = i;//定义一个标识，用来作为书签索引，书签索引从1开始
                Bookmark bm = doc.Bookmarks.get_Item(ref index);//获取遍历到的书签
                if (bm.Name == "书签" + (titleIndex + 1).ToString("00"))//判断书签的名字是否为指定格式
                {
                    result[titleIndex, 0] = bm.Start;//记录书签的开始位置
                    result[titleIndex, 1] = bm.End;//记录书签的结束位置
                    titleIndex++;//使二维数组索引加1
                }
            }
            return result;//返回记录所有书签开始和结束位置的int二维数组
        }
        #endregion

        #region 对Word文档进行拆分
        /// <summary>
        /// 对Word文档进行拆分
        /// </summary>
        /// <param name="wordPath">要拆分的Word文档路径</param>
        /// <param name="savePath">拆分后的Word存储路径</param>
        private void SplitWord(string wordPath, string savePath)
        {
            object missing = System.Reflection.Missing.Value;//作为方法传参中的缺省值
            if (!Directory.Exists(savePath))//判断Word存储路径是否存在
                Directory.CreateDirectory(savePath);//如果路径不存在，则创建
            try
            {
                //判断拆分方式：0表示按书签拆分，1表示按分页符拆分，默认为按书签拆分
                switch (flag)
                {
                    #region 按书签拆分
                    case 0:
                    default:
                        Document doc = null;//声明一个Word文档对象
                        if (doc == null)//判断Word文档对象是否为空
                            doc = ReadDocument(wordPath);//调用方法读取指定Word内容，并实例化Word文档对象
                        int[,] positions = GetPosition(doc);//获取指定Word中的所有书签的开始和结束位置
                        object oStart = 0;//定义变量，表示要拆分的开始位置
                        object oEnd = 0;//定义变量，表示要拆分的结束位置
                        int row = positions.GetLength(0);//获取二维数组的行数
                        for (int i = 0; i < row; i++)//遍历二维数组的所有行
                        {
                            if (i != row - 1)//判断是否为最后一行
                            {
                                oEnd = positions[i, 1];//如果不是最后一行，记录值，作为拆分的结束位置
                            }
                            else
                            {
                                oEnd = doc.Content.End;//如果是最后一行，则直接将Word文档的最后位置作为拆分的结束位置
                            }
                            Range tocopy = doc.Range(ref oStart, ref oEnd);//使用开始和结束位置锁定要拆分的范围
                            tocopy.Copy();//复制要拆分范围内容的所有内容
                            Document docto = CreateDocument(textBox1.Text);//调用自定义方法创建一个新的Word文档
                            docto.Content.Paste();//将复制的内容粘贴到新创建的Word文档中
                            //设置Word文档的保存路径及文件名（以编号命名）
                            object filename = savePath + "\\" + i.ToString("000") + ".docx";
                            //保存Word文档
                            docto.SaveAs(ref filename, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing);
                            docto.Close(ref missing, ref missing, ref missing);//关闭Word文档
                            oStart = oEnd;//将本次拆分的结束位置作为下一次拆分的开始位置
                        }
                        break;
                    #endregion

                    #region 按分页符拆分
                    case 1:
                        Spire.Doc.Document original = new Spire.Doc.Document();//使用Spire插件创建Word文档对象
                        original.LoadFromFile(wordPath);//加载要拆分的Word文档
                        Spire.Doc.Document newWord = new Spire.Doc.Document(textBox1.Text);//使用选择的模板创建一个新的Word文档对象
                        Spire.Doc.Section section = newWord.AddSection();//为新创建的Word文档添加一节
                        int index = 0;//根据分页来拆分文档
                        foreach (Spire.Doc.Section sec in original.Sections)//遍历文档所有节 
                        {
                            //遍历文档所有子对象  
                            foreach (Spire.Doc.DocumentObject obj in sec.Body.ChildObjects)
                            {
                                if (obj is Spire.Doc.Documents.Paragraph)//判断是否为段落
                                {
                                    //创建Word文档中的段落对象
                                    Spire.Doc.Documents.Paragraph para = obj as Spire.Doc.Documents.Paragraph;
                                    //复制并添加原有段落对象到新文档  
                                    section.Body.ChildObjects.Add(para.Clone());
                                    //遍历所有段落子对象  
                                    foreach (Spire.Doc.DocumentObject parobj in para.ChildObjects)
                                    {
                                        //判断是否为分页符
                                        if (parobj is Spire.Doc.Break && (parobj as Spire.Doc.Break).BreakType == Spire.Doc.Documents.BreakType.PageBreak)
                                        {
                                            //获取段落分页并移除，保存新文档到文件夹  
                                            int i = para.ChildObjects.IndexOf(parobj);
                                            section.Body.LastParagraph.ChildObjects.RemoveAt(i);
                                            newWord.SaveToFile(savePath + "\\" + index.ToString("000") + ".docx", Spire.Doc.FileFormat.Docx);
                                            index++;
                                            //实例化Document类对象，添加section，将原文档段落的子对象复制到新文档  
                                            newWord = new Spire.Doc.Document(textBox1.Text);
                                            section = newWord.AddSection();
                                            section.Body.ChildObjects.Add(para.Clone());
                                            if (section.Paragraphs[0].ChildObjects.Count == 0)
                                            {
                                                section.Body.ChildObjects.RemoveAt(0);//移除第一个空白段落
                                            }
                                            else
                                            {
                                                //删除分页符前的子对象
                                                while (i >= 0)
                                                {
                                                    section.Paragraphs[0].ChildObjects.RemoveAt(i);
                                                    i--;
                                                }
                                            }
                                        }
                                    }
                                }
                                //若对象为表格，则添加表格对象到新文档  
                                if (obj is Spire.Doc.Table)
                                {
                                    section.Body.ChildObjects.Add(obj.Clone());
                                }
                            }
                        }
                        //拆分后的新文档保存至指定文档  
                        newWord.SaveToFile(savePath + "\\" + index.ToString("00") + ".docx", Spire.Doc.FileFormat.Docx);
                        original.Close();
                        original.Dispose();
                        newWord.Close();
                        newWord.Dispose();
                        foreach (string file in Directory.GetFiles(savePath))//遍历拆分完的所有Word文档
                            deletePagesInFile(file, 1, 1);//删除第一页
                        break;
                        #endregion
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }//异常跳过
        }
        #endregion

        #region 删除Word中指定范围的页（因为使用Spire插件、按照分页符拆分后的Word中，第一页为分页，而且有水印信息）
        static object oMissing = System.Reflection.Missing.Value;//作为方法传参中的缺省值
        /// <summary>
        /// 删除Word中的指定页
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="objPage"></param>
        static void deletePageInFile(Document wordDoc, object objPage)
        {
            //获取打开Word文档的页数
            int pages = wordDoc.ComputeStatistics(WdStatistic.wdStatisticPages, ref oMissing);
            //设置跳转种类是按页码跳转
            object objWhat = WdGoToItem.wdGoToPage;
            //设置跳转位置是跳转到绝对位置
            object objWhich = WdGoToDirection.wdGoToAbsolute;
            //定位到Word中的指定页
            Range range1 = wordDoc.GoTo(ref objWhat, ref objWhich, ref objPage, ref oMissing);
            //定位到指定页的下一页
            Range range2 = range1.GoToNext(WdGoToItem.wdGoToPage);
            //获取要删除的开始位置
            object objStart = range1.Start;
            //获取要删除的结束位置
            object objEnd = range2.Start;
            //判断开始和结束位置是否相同
            if (range1.Start == range2.Start)
                objEnd = wordDoc.Characters.Count;//如果相同，说明是最后一页
            //记录要删除的内容文本
            string str = wordDoc.Range(ref objStart, ref objEnd).Text;
            //设置删除Word文档内容的单位为字符
            object Unit = (int)WdUnits.wdCharacter;
            object Count = 1;//标识要删除的单位数
            if (string.IsNullOrEmpty(str.Trim()))//判断要删除的内容是否为空
            {
                wordDoc.Range(ref objStart, ref objEnd).Delete(ref Unit, ref Count);//删除空行
            }
            else
            {
                wordDoc.Range(ref objStart, ref objEnd).Delete(ref Unit, ref Count);//删除指定范围的内容
            }
        }
        /// <summary>
        /// 删除Word文档中的指定范围内的页
        /// </summary>
        /// <param name="filePath">要操作的Word文档路径</param>
        /// <param name="start">开始删除的页</param>
        /// <param name="end">结束删除的页</param>
        static void deletePagesInFile(object filePath, int start, int end)
        {
            //创建Word文档对象
            Microsoft.Office.Interop.Word.Application wordApp = new ApplicationClass();
            //打开Word文档
            Document wordDoc = wordApp.Documents.Open(ref filePath, ref oMissing, ref oMissing, ref oMissing,ref oMissing, ref oMissing, ref oMissing, ref oMissing,ref oMissing, ref oMissing, ref oMissing, ref oMissing,ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            //从尾部开始向前遍历
            for (int i = end; i >= start; i--)
            {
                deletePageInFile(wordDoc, start);//调用自定义方法删除指定的页
            }
            //保存删除页码之后的Word文档
            wordDoc.SaveAs(ref filePath,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing);
            wordDoc.Close(ref oMissing, ref oMissing, ref oMissing);//关闭Word文档
        }
        #endregion

        //执行Word拆分操作
        private void Button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")//判断是否选择了模板文件
            {
                MessageBox.Show("请选择模板文件！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                button1.Focus();//为选择模板文件的按钮设置焦点
            }
            else if (textBox2.Text == "")//判断是否选择了要拆分的Word文档路径
            {
                MessageBox.Show("请选择要拆分的Word所在路径！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                button2.Focus();//为选择要拆分Word文档路径的按钮设置焦点
            }
            else if (textBox3.Text == "")//判断是否选择了拆分后文件的存储路径
            {
                MessageBox.Show("请选择拆分后文件的存储路径！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                button3.Focus();//为选择拆分后文件存储路径的按钮设置焦点
            }
            else
            {
                string[] files = Directory.GetFiles(textBox2.Text);//获取要拆分的所有Word文档
                int count = files.Length;//获取文件个数，为了设置进度条的显示
                progressBar1.Value = 0;//设置进度条初始值为0
                progressBar1.Maximum = count;//设置进度条最大值为要操作文件的个数
                ThreadPool.QueueUserWorkItem(//开始线程池，防止窗体出现“假死”状态
                (pp) =>//使用lambda表达式
                {
                    button1.Enabled = false;
                    for (int i = 0; i < count; i++)//遍历文件
                    {
                        FileInfo file = new FileInfo(files[i]);//使用遍历到的文件创建文件对象
                        //筛选Word文件，并屏蔽缓存文件
                        if ((file.Extension == ".docx" || file.Extension == ".doc") && !file.Name.Contains("$"))
                            //使用自定义的SplitWord方法对Word文档进行拆分 
                            SplitWord(files[i], textBox3.Text.Trim(new char[] { '\\' }) + "\\" + Path.GetFileNameWithoutExtension(files[i]));
                        progressBar1.Value += 1;//设置进度条的值加1
                    }
                    //成功提示
                    MessageBox.Show("拆分成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    button1.Enabled = true;
                    try
                    {
                        //关闭遗留的Word进程
                        Process[] MyProcess = Process.GetProcessesByName("Microsoft Word");
                        MyProcess[0].Kill();
                    }
                    catch { }
                });
            }
        }

        //选择模板文件
        private void Button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();//创建打开文件对话框
            openfile.Filter = "Word模板文件|*.dotx;*.dot";//设置只能打开Word模板（.dotx或者.dot格式）
            if (openfile.ShowDialog() == DialogResult.OK)//判断是否选择了文件
            {
                textBox1.Text = openfile.FileName;//显示选择的Word模板
            }
        }

        //选择Word文档所在路径
        private void Button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dir = new FolderBrowserDialog();//创建浏览文件对话框
            if (dir.ShowDialog() == DialogResult.OK)//判断是否选择了路径
            {
                textBox2.Text = dir.SelectedPath;//显示选择的Word文档所在路径
            }
        }

        //打开保存路径进行查看
        private void Button5_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(textBox3.Text);//打开保存路径进行查看
        }
       
        //记录保存路径
        private void Button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dir = new FolderBrowserDialog();//创建浏览文件对话框
            if (dir.ShowDialog() == DialogResult.OK)//判断是否选择了路径
            {
                textBox3.Text = dir.SelectedPath;//显示选择的路径
            }
        }

        //选择拆分方式
        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            flag = radioButton1.Checked ? 0 : 1;//标识按书签拆分还是分页符拆分
        }

        //退出当前应用程序
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();//关闭当前应用程序
        }

        //“整理命名”超链接单击事件
        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new Form2().ShowDialog();//打开文件重命名窗体
        }

        //关闭遗留的Word进程
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (Process p in Process.GetProcesses()) //遍历所有进程
            {
                if (p.ProcessName.ToUpper().Contains("WINWORD"))//判断是否为Word进程
                {
                    try
                    {
                        p.Kill();//关闭Word进程
                        p.WaitForExit();//等待进程退出
                    }
                    catch { }//异常通过
                }
            }
        }
    }
}