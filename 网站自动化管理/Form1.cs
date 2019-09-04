using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CefSharp;
using CefSharp.WinForms;
using CefSharp.WinForms.Internals;
using Microsoft.VisualBasic;
using 网站自动化管理.Controls;
using 网站自动化管理.cs;
using static 网站自动化管理.cs.cmds;

namespace 网站自动化管理
{
    public partial class Form1 : Form
    {
        private readonly ChromiumWebBrowser browser;

        private IList<string> codeList = null;
        private int codeIndex = 0;
        private const int RUN_WAIT = 100;
        public Form1()
        {
            InitializeComponent();


            Text = "CefSharp";
            WindowState = FormWindowState.Maximized;

            // 避免页面加载不出来, 建议加上这句
            Cef.Initialize(new CefSettings());
            browser = new ChromiumWebBrowser("https://www.baidu.com/")
            {
                Dock = DockStyle.Fill,
                LifeSpanHandler = new OpenPageSelf()

            };

            splitContainer1.Panel2.Controls.Add(browser);


            browser.IsBrowserInitializedChanged += OnIsBrowserInitializedChanged;
            browser.LoadingStateChanged += OnLoadingStateChanged;
            browser.ConsoleMessage += OnBrowserConsoleMessage;
            browser.StatusMessage += OnBrowserStatusMessage;
            browser.TitleChanged += OnBrowserTitleChanged;
            browser.AddressChanged += OnBrowserAddressChanged;
            browser.FrameLoadEnd += Browser_FrameLoadEnd;

            var bitness = Environment.Is64BitProcess ? "x64" : "x86";
            var version = String.Format("Chromium: {0}, CEF: {1}, CefSharp: {2}, Environment: {3}", Cef.ChromiumVersion, Cef.CefVersion, Cef.CefSharpVersion, bitness);
            DisplayOutput(version);
        }

        private async void Browser_FrameLoadEnd(object sender, FrameLoadEndEventArgs e)
        {
            // throw new NotImplementedException();
            //一个网页会调用多次,需要手动自己处理逻辑
            var url = e.Url;
            var result = await browser.GetSourceAsync();
            var html = result;

            ////调用js
            // toolStripButton10_Click(null, null);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cmds.browser = browser;
            timer1.Interval =RUN_WAIT;
        }

        private void OnIsBrowserInitializedChanged(object sender, EventArgs e)
        {
            var b = ((ChromiumWebBrowser)sender);

            this.InvokeOnUiThreadIfRequired(() => b.Focus());

        }

        private void OnBrowserConsoleMessage(object sender, ConsoleMessageEventArgs args)
        {
            DisplayOutput(string.Format("Line: {0}, Source: {1}, Message: {2}", args.Line, args.Source, args.Message));
        }

        private void OnBrowserStatusMessage(object sender, StatusMessageEventArgs args)
        {
            // this.InvokeOnUiThreadIfRequired(() => statusLabel.Text = args.Value);
        }

        private void OnLoadingStateChanged(object sender, LoadingStateChangedEventArgs args)
        {
            SetCanGoBack(args.CanGoBack);
            SetCanGoForward(args.CanGoForward);

            this.InvokeOnUiThreadIfRequired(() => SetIsLoading(!args.CanReload));
        }

        private void OnBrowserTitleChanged(object sender, TitleChangedEventArgs args)
        {
            this.InvokeOnUiThreadIfRequired(() => Text = args.Title);
        }

        private void OnBrowserAddressChanged(object sender, AddressChangedEventArgs args)
        {
            // this.InvokeOnUiThreadIfRequired(() => urlTextBox.Text = args.Address);
        }

        private void SetCanGoBack(bool canGoBack)
        {
            //  this.InvokeOnUiThreadIfRequired(() => backButton.Enabled = canGoBack);
        }

        private void SetCanGoForward(bool canGoForward)
        {
            // this.InvokeOnUiThreadIfRequired(() => forwardButton.Enabled = canGoForward);
        }

        private void SetIsLoading(bool isLoading)
        {
            //goButton.Text = isLoading ?
            //    "Stop" :
            //    "Go";
            //goButton.Image = isLoading ?
            //    Properties.Resources.nav_plain_red :
            //    Properties.Resources.nav_plain_green;

            HandleToolStripLayout();
        }

        public void DisplayOutput(string output)
        {
            //  this.InvokeOnUiThreadIfRequired(() => outputLabel.Text = output);
        }

        private void HandleToolStripLayout(object sender, LayoutEventArgs e)
        {
            HandleToolStripLayout();
        }

        private void HandleToolStripLayout()
        {
            //var width = toolStrip1.Width;
            //foreach (ToolStripItem item in toolStrip1.Items)
            //{
            //    if (item != urlTextBox)
            //    {
            //        width -= item.Width - item.Margin.Horizontal;
            //    }
            //}
            //urlTextBox.Width = Math.Max(0, width - urlTextBox.Margin.Horizontal - 18);
        }

        private void ExitMenuItemClick(object sender, EventArgs e)
        {
            browser.Dispose();
            Cef.Shutdown();
            Close();
        }

        private void GoButtonClick(object sender, EventArgs e)
        {
            // LoadUrl(urlTextBox.Text);
        }

        private void BackButtonClick(object sender, EventArgs e)
        {
            browser.Back();
        }

        private void ForwardButtonClick(object sender, EventArgs e)
        {
            browser.Forward();
        }

        private void UrlTextBoxKeyUp(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode != Keys.Enter)
            //{
            //    return;
            //}

            //LoadUrl(urlTextBox.Text);
        }

        private void LoadUrl(string url)
        {
            if (Uri.IsWellFormedUriString(url, UriKind.RelativeOrAbsolute))
            {
                browser.Load(url);
            }
        }


        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            browser.Focus();


            Task.Run(() =>//多线程
            {
                // ShowLabel("task run start");

                int x = browser.Width / 2;
                int y = browser.Height / 2;
                y -= 50;
                y -= 50;
                x -= 30;

                int x2 = x + 50;
                int y2 = y + 80;
                string msg = string.Format("x1={0},y1={1},x2={2},y2={3},W={4},H={5}", x, y, x2, y2, browser.Width, browser.Height);
                //ShowLabel(msg);
                //MessageBox.Show(msg);
                CefMouseDrawLine(x, y, x2, y2);//第一条线

                x = x2;
                y = y2;
                x2 = x + 100;
                y2 = y - 20;
                CefMouseDrawLine(x, y, x2, y2);//第二条线
            });
        }



        //画线。鼠标左键按下，移动，移动...左键抬起。
        public void CefMouseDrawLine(int x1, int y1, int x2, int y2)
        {
            //Task.Run(() =>
            //{
            // ShowLabel("------------start");
            var host = browser.GetBrowser().GetHost();

            host.SendMouseClickEvent(x1, y1, MouseButtonType.Left, false, 1, CefEventFlags.None);//按下鼠标左键
            Thread.Sleep(3);
            double x = x1;
            double y = y1;
            int xlen = x2 - x1;
            int ylen = y2 - y1;
            double xs = 1;
            double ys = 1;

            int z = (int)Math.Sqrt(xlen * xlen + ylen * ylen);
            xs = (double)xlen / (double)z;
            ys = (double)ylen / (double)z;

            for (int i = 1; i < z; i++)
            {
                x = (x + xs);
                y = (y + ys);
                //ShowLabel("x=" + x + ",y=" + y);
                Thread.Sleep(3);
                host.SendMouseMoveEvent((int)x, (int)y, false, CefEventFlags.LeftMouseButton);//移动鼠标
            }
            Thread.Sleep(3);
            host.SendMouseClickEvent(x2, y2, MouseButtonType.Left, true, 1, CefEventFlags.None);//抬起鼠标左键

            //ShowLabel("------------end");
            //});
        }

        //退出
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            browser.Dispose();
            Cef.Shutdown();
            Close();
        }

        //工具
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            browser.ShowDevTools();
        }



        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Cef.Shutdown();
        }





        private void toolStripButton10_Click(object sender, EventArgs e)
        {

            Thread.Sleep(10000);
            cmds.clicklink("tj_login");
            Thread.Sleep(2000);
            cmds.clickDom("TANGRAM__PSP_10__footerULoginBtn");
            Thread.Sleep(2000);
            cmds.inputtxtByid("TANGRAM__PSP_10__userName", "13377864308");
            Thread.Sleep(2000);
            cmds.inputtxtByid("TANGRAM__PSP_10__password", "11111111a.");
            Thread.Sleep(6000);
            //submit("TANGRAM__PSP_10__form");
            cmds.clickDom("TANGRAM__PSP_10__submit");
            Thread.Sleep(2000);
            string a = cmds.getTxt("user-name", false);

        }

        private void 登陆链接ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cmds.clicklink("login");
            Thread.Sleep(1000);
            cmds.inputtxtByid("login_name_d", "openmmoo@163.com");
            Thread.Sleep(500);
            cmds.inputtxtByid("login_pass_d", "thisisadog");
            Thread.Sleep(500);
            //submit("TANGRAM__PSP_10__form");
            cmds.clicklink("login_button", false);
            Thread.Sleep(10000);
            cmds.clicklink("SG_Publish", false);
            Thread.Sleep(1000);
            cmds.inputtxtByid("articleTitle", "测试一下标题");
            //inputtxtByid("SinaEditor_Iframe", "thisisadog");
            Thread.Sleep(1500);
            cmds.clickDom("SinaEditor_Iframe", 100, 100);
            Thread.Sleep(1500);
            cmds.cmdInput("测试一下，剪切板的使用");
            Thread.Sleep(1500);
            cmds.select("componentSelect", "3");
            Thread.Sleep(500);
            cmds.clickDomBycsname("cTit", 2, 100);
            Thread.Sleep(1500);
            cmds.cmdInput("剪切板的使用，是win提供的一个非常有用的东西" +
                "没有什么不可以用剪切板完成！" +
                "但也有不好的，容易复制，造成假文章到处流传" +
                "一定要防止");
            Thread.Sleep(2000);
            cmds.clickDomBycsname("cTit", 2);
            Thread.Sleep(2000);
            cmds.clickDomBycsname("editBtn2");

        }


        string oldText = "";
        private void linkhelp_Click(object sender, EventArgs e)
        {
            if (linkhelp.Text == "帮助")
            {

                oldText = txtCode.Text;
                linkhelp.Text = "关闭";
                txtCode.Text = @"========================
按键助手  v1.0 
========================
作者：山树有约
日期：2018-12-31

输入格式为：命令:内容 每行一条命令

INPUT         输入文本
RUN           运行程序
KEY           模拟按键
SLEEP         延时
MOUSE_MOVE    鼠标移动
MOUSE_CLICK   鼠标单击
MOUSE_DBCLICK 鼠标双击
SCREEN        窗口截屏
ALL_SCREEN    全屏截屏

KEY 辅助说明 按照C# SendKeys.Send 函数要求

字母或数字 a-z A-Z 0-9
Alt    %
Ctrl   ^
Shift  +
向上键 {UP} 
向下键 {DOWN}  
向左键 {LEFT}  
向右键 {RIGHT}  
Enter  {ENTER} 或 ~  
Backspace {BACKSPACE}、{BS} 或 {BKSP}  
Break     {BREAK}  
Caps Lock {CAPSLOCK}  
Scroll Lock   {SCROLLLOCK}  
Print Screen  {PRTSC}（保留供将来使用）  
Del 或 Delete {DELETE} 或 {DEL}  
End   {END}  
Esc   {ESC}  
Help  {HELP}  
Home  {HOME}  
Ins 或 Insert  {INSERT} 或 {INS}  
Num Lock   {NUMLOCK}  
Page Down  {PGDN}  
Page Up    {PGUP}  
Tab        {TAB}  
F1-F16     {F1-F16} 
数字键盘加号 {ADD}  
数字键盘减号 {SUBTRACT} 
数字键盘乘号 {MULTIPLY}  
数字键盘除号 {DIVIDE} 
特殊键 {{} {%}
重复键 {h 10}
组合键 ^(AC)
";
            }
            else
            {
                txtCode.Text = oldText;
                linkhelp.Text = "帮助";
                txtCode.Focus();
            }
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            if (linkhelp.Text == "关闭")
            {
                txtCode.Text = oldText;
                linkhelp.Text = "帮助";
            }
            if (btnGo.Text == "停 止")
            {
                timer1.Stop();
                btnGo.Text = "开 始";
            }
            else if (btnGo.Text == "开 始")
            {

                if (txtCode.Text == "")
                {
                    MessageBox.Show("没有脚本");
                    return;
                }
                //自动执行

                codeList = new List<string>();
                codeIndex = 0;
                foreach (string r in txtCode.Text.Split('\r'))
                {
                    if (r.Trim() != "")
                    {
                        codeList.Add(r.Trim());
                    }
                }
                //LoadData();
                btnGo.Text = "停 止";
                timer1.Interval = RUN_WAIT;
                timer1.Start();
            }
            else if (btnGo.Text == "暂 停")
            {
                btnGo.Text = "停 止";
                runflay = 1;
                //timer1.Start();
            }

        }
        int runflay = 1;
        long starttime;
        long datediff = 0;
        //控件代码运行
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (runflay == 1)
            {
                if (codeList.Count <= 0)
                {
                    MessageBox.Show("没有脚本。");
                    timer1.Stop();
                    btnGo.Text = "开 始";
                    return;
                }
                if (codeIndex >= codeList.Count)
                {
                    MessageBox.Show("运行完毕");
                    timer1.Stop();
                    btnGo.Text = "开 始";
                    return;
                }
                lblInfo.Text = string.Format("{0}/{1} {2}", codeIndex + 1, codeList.Count, codeList[codeIndex]);
                string code = codeList[codeIndex];
                int split = code.IndexOf(':');
                string codeType = "";
                string codeContent = "";
                if (split > 0)
                {
                    codeType = code.Substring(0, split);
                    codeContent = code.Substring(split + 1);
                }
                else
                {
                    codeType = code;
                }
                codeType = codeType.Trim().ToUpper();
                if (Enum.IsDefined(typeof(CmdType), codeType))
                {
                    try
                    {
                        CmdType cmdType = (CmdType)Enum.Parse(typeof(CmdType), codeType);
                        switch (cmdType)
                        {
                            case CmdType.INPUT:

                                cmdInput(codeContent);

                                runflay = 2;
                                break;
                            case CmdType.RUN:
                                Task.Run(() =>//多线程
                                {
                                   string s= cmdRun(codeContent);
                                    this.InvokeOnUiThreadIfRequired(() => label1.Text = s);
                                });
                                runflay = 2;
                                break;
                            case CmdType.KEY:
                                Task.Run(() =>//多线程
                                {
                                    string s = cmdKey(codeContent);
                                    this.InvokeOnUiThreadIfRequired(() => label1.Text = s);
                                });
                                runflay = 2;
                                break;
                            case CmdType.SLEEP:
                                
                                //datediff = cmdSleep(codeContent);
                                if (long.TryParse(codeContent, out datediff))
                                {
                                    starttime = DateTime.Now.Ticks;
                                    runflay = 3;
                                }else
                                {
                                    runflay = -1;
                                    btnGo.Text = "暂 停";
                                    //timer1.Stop();
                                    this.Text = "错误！";
                                }
                               
                                break;
                            case CmdType.MOUSE_MOVE:
                                Task.Run(() =>//多线程
                                {
                                    string s = cmdMouseMove(codeContent);
                                    this.InvokeOnUiThreadIfRequired(() => label1.Text = s);
                                });
                                runflay = 2;
                                break;
                            case CmdType.MOUSE_CLICK:
                                Task.Run(() =>//多线程
                                {
                                    string s = cmdMouseClick(codeContent);
                                    this.InvokeOnUiThreadIfRequired(() => label1.Text = s);
                                });
                                runflay = 2;
                                break;
                            case CmdType.MOUSE_DBCLICK:
                                Task.Run(() =>//多线程
                                {
                                    string s = cmdMouseDBClick(codeContent);
                                    this.InvokeOnUiThreadIfRequired(() => label1.Text = s);
                                });
                                runflay = 2;
                                break;
                            case CmdType.SCREEN:
                                //cmdScreen();
                                break;
                            case CmdType.ALL_SCREEN:
                                //cmdAllScreen();
                                break;
                            case CmdType.DOCOLOR:
                                //cmdDoColor(codeContent);
                                break;
                            case CmdType.IFCOLOR:
                                //cmdifColor(codeContent);
                                break;
                            case CmdType.ELSE:
                                //cmdElse(codeContent);
                                break;
                            case CmdType.GOTO:
                                //cmdGoto(codeContent);
                                break;
                            case CmdType.END:
                                //codeIndex = 100000000;
                                break;
                            case CmdType.PLAY:
                                cmdPlay();
                                break;
                            case CmdType.SLEEPCLICK:
                                //cmdSleepClick();
                                break;
                            default:
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        btnGo.PerformClick();
                        MessageBox.Show("运行[" + code + "]时失败！\r\n\r\n错误原因：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else if (runflay == 2)
            {
                if (label1.Text == "1")
                {
                    codeIndex++;
                    if (codeIndex >= codeList.Count)
                    {
                        lblInfo.Text = "运行结束";
                        timer1.Stop();
                        btnGo.Text = "开 始";
                    }

                    runflay = 1;
                }

            }
            else if (runflay == 3)
            {
                long lsdate = (DateTime.Now.Ticks - starttime)/10000; //DateDiff(starttime, DateTime.Now);
                Console.WriteLine(lsdate.ToString());
                if (lsdate >= datediff) { 
                    codeIndex++;
                    if (codeIndex >= codeList.Count)
                    {
                        lblInfo.Text = "运行结束";
                        timer1.Stop();
                        btnGo.Text = "开 始";
                    }
                    runflay = 1;
                }
            }
        }


       

        /// <summary>
        /// C# 求两个时间差值（时，分，秒等等）
        /// </summary>
        /// <param name="DateTime1"></param>
        /// <param name="DateTime2"></param>
        /// <returns></returns>
        public int DateDiff(DateTime DateTime1, DateTime DateTime2)
        {
            //string dateDiff = null;
            TimeSpan ts1 = new TimeSpan(DateTime1.Ticks);
            TimeSpan ts2 = new TimeSpan(DateTime2.Ticks);
            TimeSpan ts = ts1.Subtract(ts2).Duration();
            //dateDiff = ts.Days.ToString() + "天" + ts.Hours.ToString() + "小时" + ts.Minutes.ToString() + "分钟" + ts.Seconds.ToString() + "秒";
            return (DateTime2 - DateTime1).Milliseconds;
        }

        private void lblInfo_Click(object sender, EventArgs e)
        {
            if (codeIndex >= codeList.Count)
            {
                MessageBox.Show("命令不在范围内");
                return;
            }
            string s1= string.Format("{0}/{1} {2}", codeIndex + 1, codeList.Count, codeList[codeIndex]);
            string s= Interaction.InputBox(s1, "修改命令", codeList[codeIndex], 100, 200); 
            if(s!= codeList[codeIndex])
            {
                codeList[codeIndex] = s;
            }

        }
    }
}
