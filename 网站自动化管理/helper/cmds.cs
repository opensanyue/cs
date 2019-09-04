using CefSharp;
using CefSharp.WinForms;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 网站自动化管理.cs
{
    public static class cmds
    {
        public static  ChromiumWebBrowser browser;
        

        /// <summary>
        ///  单击链接
        /// </summary>
        /// <param name="linkname">链接类名或id名</param>
        /// <param name="isName">为真是类名，否则为id名，默认为真</param>
        public static string clicklink(string linkname, Boolean isName = true)
        {
            if (browser == null) return null;
            var frame = browser.GetMainFrame();
            string linkstr = "";
            if (isName == true)
            {
                linkstr = string.Format("document.getElementsByClassName('{0}')[0].click();", linkname);
            }
            else
            {
                linkstr = string.Format("document.getElementById('{0}').click();", linkname);
            }

            Task<CefSharp.JavascriptResponse> t1 = frame.EvaluateScriptAsync(linkstr, null);
            t1.Wait();
            // t.Result 是 CefSharp.JavascriptResponse 对象
            // t.Result.Result 是一个 object 对象，来自js的 callTest2() 方法的返回值
            if (t1.Result != null)
                if (t1.Result.Result != null)
                {
                    string a = t1.Result.Result.ToString();
                    return a;
                }

            return null;
        }


        /// <summary>
        /// 根据id或类名获取元素内容
        /// </summary>
        /// <param name="domid">id或类名</param>
        /// <param name="isid">指定当前是id或类名,默认是id</param>
        /// <returns></returns>
        public static string getTxt(string domid, Boolean isid = true)
        {
            if (browser == null) return null;
            var frame = browser.GetMainFrame();
            string linkstr = "";
            if (isid == true)
            {
                linkstr = string.Format("(function() {{return document.getElementById('{0}').innerText; }})();", domid);
            }
            else
            {
                //linkstr = string.Format("(function() {{return document.getElementsByName('{0}')[0].value; }})();", domid);
                linkstr = string.Format("(function() {{return document.getElementsByClassName('{0}')[0].innerText; }})();", domid);
            }
            Task<CefSharp.JavascriptResponse> t1 = frame.EvaluateScriptAsync(linkstr, null);
            t1.Wait();
            // t.Result 是 CefSharp.JavascriptResponse 对象
            // t.Result.Result 是一个 object 对象，来自js的 callTest2() 方法的返回值
            if (t1.Result != null)
                if (t1.Result.Result != null)
                {
                    string a = t1.Result.Result.ToString();
                    return a;
                }

            return "";
        }

        /// <summary>
        /// 为文本框设置文本
        /// </summary>
        /// <param name="txtid">文本框id</param>
        /// <param name="txt">添加内容</param>
        public static void inputtxtByid(string txtid, string txt)
        {
            if (browser == null) return ;
            var frame = browser.GetMainFrame();
            string linkstr = string.Format(@"document.getElementById('{0}').value ='{1}';", txtid, txt);
            // var task = frame.EvaluateScriptAsync(@" var form = document.getElementById('form');form.submit();", null);
            Task<CefSharp.JavascriptResponse> task = frame.EvaluateScriptAsync(linkstr, null);

        }


        /// <summary>
        /// 提交文本到服务器
        /// </summary>
        /// <param name="formid">form的id</param>
        public static void submit(string formid)
        {
            if (browser == null) return ;
            var frame = browser.GetMainFrame();
            string linkstr = string.Format(@"var form = document.getElementById('{0}');form.submit();", formid);
            var task = frame.EvaluateScriptAsync(linkstr, null);
        }

        public static void select(string selectid, string txt)
        {
            if (browser == null) return ;
            var frame = browser.GetMainFrame();
            string linkstr = string.Format(@"function set_select_checked(selectId, checkValue){{  
                                                    var select = document.getElementById(selectId);  
 
                                                    for (var i = 0; i < select.options.length; i++){{  
                                                        if (select.options[i].value == checkValue){{  
                                                            select.options[i].selected = true;  
                                                            break;  
                                                        }} 
                                                    }}
                                                  
                                                }}
                                                
                                               set_select_checked('{0}', {1});
                                            ", selectid, txt);

            var task = frame.EvaluateScriptAsync(linkstr, null);
        }

        public static void clickDomBycsname(string domid, int index = 0, int leftoff = 10, int topoff = 10)
        {
            if (browser == null) return ;
            var frame = browser.GetMainFrame();

            string linkstr = string.Format(@"(function() {{ var el = document.getElementsByClassName('{0}')[{1}]; el = el.getBoundingClientRect(); return (el.left + window.scrollX)+ ','+(el.top + window.scrollY)   }})();", domid, index); //string.Format("document.getElementsByName('{0}')[0].click();", linkname);
            Task<CefSharp.JavascriptResponse> t1 = frame.EvaluateScriptAsync(linkstr);
            t1.Wait();
            // t.Result 是 CefSharp.JavascriptResponse 对象
            // t.Result.Result 是一个 object 对象，来自js的 callTest2() 方法的返回值
            if (t1.Result != null)
                if (t1.Result.Result != null)
                {
                    string a = t1.Result.Result.ToString();
                    String[] b = a.Split(',');
                    if (b != null)
                        if (b.Length == 2)
                        {
                            float left = Convert.ToSingle(b[0]) + leftoff;
                            float top = Convert.ToSingle(b[1]) + topoff;
                            //this.InvokeOnUiThreadIfRequired(() => this.Text = left + " : " + top);
                            var host = browser.GetBrowser().GetHost();
                            Thread.Sleep(3);
                            host.SendMouseClickEvent((int)left, (int)top, MouseButtonType.Left, false, 1, CefEventFlags.None);//按下鼠标左键
                            Thread.Sleep(3);
                            host.SendMouseMoveEvent((int)left, (int)top, false, CefEventFlags.LeftMouseButton);//移动鼠标
                            Thread.Sleep(3);
                            host.SendMouseClickEvent((int)left, (int)top, MouseButtonType.Left, true, 1, CefEventFlags.None);//抬起鼠标左键
                            Thread.Sleep(3);
                        }
                }
        }


        public static void clickDom(string domid, int leftoff = 10, int topoff = 10)
        {
            if (browser == null) return ;
            var frame = browser.GetMainFrame();

            string linkstr = string.Format(@"(function() {{ var el = document.getElementById('{0}'); el = el.getBoundingClientRect(); return (el.left + window.scrollX)+ ','+(el.top + window.scrollY)   }})();", domid); //string.Format("document.getElementsByName('{0}')[0].click();", linkname);
            Task<CefSharp.JavascriptResponse> t1 = frame.EvaluateScriptAsync(linkstr);
            t1.Wait();
            // t.Result 是 CefSharp.JavascriptResponse 对象
            // t.Result.Result 是一个 object 对象，来自js的 callTest2() 方法的返回值
            if (t1.Result != null)
                if (t1.Result.Result != null)
                {
                    string a = t1.Result.Result.ToString();
                    String[] b = a.Split(',');
                    if (b != null)
                        if (b.Length == 2)
                        {
                            float left = Convert.ToSingle(b[0]) + leftoff;
                            float top = Convert.ToSingle(b[1]) + topoff;
                            //this.InvokeOnUiThreadIfRequired(() => this.Text = left + " : " + top);
                            var host = browser.GetBrowser().GetHost();
                            Thread.Sleep(3);
                            host.SendMouseClickEvent((int)left, (int)top, MouseButtonType.Left, false, 1, CefEventFlags.None);//按下鼠标左键
                            Thread.Sleep(3);
                            host.SendMouseMoveEvent((int)left, (int)top, false, CefEventFlags.LeftMouseButton);//移动鼠标
                            Thread.Sleep(3);
                            host.SendMouseClickEvent((int)left, (int)top, MouseButtonType.Left, true, 1, CefEventFlags.None);//抬起鼠标左键
                            Thread.Sleep(3);
                        }
                }
        }

        #region 基本命令
        public static string cmdInput(string str)
        {
            Clipboard.Clear();
            Clipboard.SetText(str);
            SendKeys.SendWait("^v");
            //this.InvokeOnUiThreadIfRequired(() => label1.Text = "1");
            return "1";
        }


        public static string cmdRun(string str)
        {
            Process.Start(str);
            // this.InvokeOnUiThreadIfRequired(() => label1.Text = "1");
            return "1";
        }
        public static string cmdKey(string str)
        {
            SendKeys.SendWait(str);
            //this.InvokeOnUiThreadIfRequired(() => label1.Text = "1");
            return "1";
        }

       
        //public static string cmdSleep(string str)
        //{
        //    int t = 0;
        //    if (int.TryParse(str, out t))
        //    {
        //        //if (t > 100)
        //        //{
        //        //    timer1.Interval = t;
        //        //}
        //        return t;
        //    }
        //    return 0;
        //}

        //private void cmdSleep(string str)
        //{
        //    int t = 0;
        //    if (int.TryParse(str, out t))
        //    {
        //        Thread.Sleep(t);
        //        this.InvokeOnUiThreadIfRequired(() => label1.Text = "1");
        //    }
        //}




        [DllImport("winmm.dll")]
        public static extern bool PlaySound(string pszSound, int hmod, int fdwSound);//播放windows音乐，重载
        public const int SND_FILENAME = 0x00020000;
        public const int SND_ASYNC = 0x0001;
        public static string cmdPlay()
        {
            string strPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            strPath = strPath + "\\sound\\Audio.wav";
            PlaySound(strPath, 0, SND_ASYNC | SND_FILENAME);//播放音乐
            return "1";
        }



        public static string cmdScreen()
        {
            Clipboard.Clear();
            SendKeys.SendWait("{PRTSC}");
            if (Clipboard.ContainsImage())
            {
                Image img = Clipboard.GetImage();
                savePic(img);
                img.Dispose();
                Clipboard.Clear();
            }
            return "1";
        }
        public static string cmdAllScreen()
        {
            Rectangle rect = Screen.PrimaryScreen.Bounds;
            using (Bitmap bmp = new Bitmap(rect.Width, rect.Height))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(0, 0, 0, 0, new Size(rect.Width, rect.Height));
                }
                savePic(bmp);
            }
            return "1";
        }
        public static string savePic(Image img)
        {
            string strScreenPath = Application.StartupPath + "\\SCREEN\\";
            if (!System.IO.Directory.Exists(strScreenPath))
            {
                System.IO.Directory.CreateDirectory(strScreenPath);
            }
            img.Save(strScreenPath + Guid.NewGuid().ToString() + ".png", System.Drawing.Imaging.ImageFormat.Png);
            return "1";
        }

        [Flags]
        enum MouseEventFlag : uint
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            MiddleDown = 0x0020,
            MiddleUp = 0x0040,
            XDown = 0x0080,
            XUp = 0x0100,
            Wheel = 0x0800,
            VirtualDesk = 0x4000,
            Absolute = 0x8000,
        }

        [DllImport("user32.dll")]
        static extern void mouse_event(MouseEventFlag flags, int dx, int dy, uint data, UIntPtr ext);

        public static string cmdMouseMove(string str)
        {
            if (str.Trim() == "") return null;
            string[] strs = str.Split(',');
            if (strs.Length == 2)
            {
                int x = int.Parse(strs[0]);
                int y = int.Parse(strs[1]);
                Cursor.Position = new Point(x, y);
            }
            //this.InvokeOnUiThreadIfRequired(() => label1.Text = "1");
            return "1";
        }


        public static string cmdMouseClick(string str)
        {
            if (str.Trim() != "")
                cmdMouseMove(str);

            mouse_event(MouseEventFlag.LeftDown, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MouseEventFlag.LeftUp, 0, 0, 0, UIntPtr.Zero);
            //this.InvokeOnUiThreadIfRequired(() => label1.Text = "1");
            return "1";
        }
        public static string cmdMouseDBClick(string str)
        {
            cmdMouseClick(str);
            cmdMouseClick("");
            //this.InvokeOnUiThreadIfRequired(() => label1.Text = "1");
            return "1";
        }
        #endregion


        #region 命令类型
        /// <summary>
        /// 命令类型
        /// </summary>
        public enum CmdType
        {
            /// <summary>
            /// 输入
            /// </summary>
            INPUT,

            /// <summary>
            /// 运行
            /// </summary>
            RUN,

            /// <summary>
            /// 按键
            /// </summary>
            KEY,

            /// <summary>
            /// 暂停
            /// </summary>
            SLEEP,

            /// <summary>
            /// 鼠标移动
            /// </summary>
            MOUSE_MOVE,

            /// <summary>
            /// 鼠标单击
            /// </summary>
            MOUSE_CLICK,

            /// <summary>
            /// 鼠标双击
            /// </summary>
            MOUSE_DBCLICK,

            /// <summary>
            /// 窗口截屏
            /// </summary>
            SCREEN,

            /// <summary>
            /// 全屏截屏
            /// </summary>
            ALL_SCREEN,

            /// <summary>
			/// 
			/// </summary>
			IFCOLOR,

            /// <summary>
            /// 
            /// </summary>
            ELSE,

            /// <summary>
            /// 
            /// </summary>
            ENDIF,

            /// <summary>
            /// 循环
            /// </summary>
            GOTO,

            /// <summary>
            /// 检查色
            /// </summary>
            DOCOLOR,

            /// <summary>
            /// 退出
            /// </summary>
            END,


            /// <summary>
            /// 放声音
            /// </summary>
            PLAY,
            /// <summary>
            /// 暂停（点击）
            /// </summary>
            SLEEPCLICK

        }
        #endregion
    }
}
