using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
//using Autodesk.AutoCAD.Interop;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

[assembly: ExtensionApplication(typeof(ExcelForm.IExtension))]
[assembly: CommandClass(typeof(ExcelForm.IExtension))]

namespace ExcelForm
{

    public class IExtension : IExtensionApplication
    {
        /// <summary>
        /// 加载程序时的初始化操作
        /// </summary>
        void IExtensionApplication.Initialize()
        {
            Autodesk.AutoCAD.Interop.AcadApplication cadapp = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication as Autodesk.AutoCAD.Interop.AcadApplication;
            for (int i = 0; i < cadapp.MenuGroups.Count; i++)
            {
                if ("DCTABLE".Equals(cadapp.MenuGroups.Item(i).Name, StringComparison.OrdinalIgnoreCase))
                {
                    cadapp.MenuGroups.Item(i).Unload();
                    break;
                }
            }
            Autodesk.AutoCAD.ApplicationServices.Application.BeginQuit += Application_BeginQuit;
            initMenu(); 
        }

        /// <summary>
        /// 即将退出事件处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Application_BeginQuit(object sender, EventArgs e)
        {
            try
            {
                string cuiPath = GetPath() + "\\dctable.CUIX";
                string newCuiPath = cuiPath.Split('\\')[0] + "\\TempTableCAD\\dctable.CUIX";
                File.Delete(newCuiPath);
            }
            catch(System.Exception ex)
            {

            }
        }

        /// <summary>
        ///  初始化菜单 以及界面
        /// </summary>
        public void initMenu()
        {
            try
            {
                Autodesk.AutoCAD.Interop.AcadApplication cadapp = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication as Autodesk.AutoCAD.Interop.AcadApplication;
                var doc = cadapp.ActiveDocument;
                string cuiPath = GetPath() + "\\dctable.CUIX";
                if (File.Exists(cuiPath))
                {
                    string newCuiPath = cuiPath.Split('\\')[0]+ "\\TempTableCAD\\dctable.CUIX";
                    if(!Directory.Exists(Path.GetDirectoryName(newCuiPath)))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(newCuiPath));
                    }
                    File.Copy(cuiPath, newCuiPath, true);
                    //string commandNetload = string.Format("(command \"cuiload\" \"{0}\" ) ", newCuiPath);
                    string commandNetload = string.Format("(command \"cuiload\" ) ");
                    doc.SendCommand("(command \"filedia\" 0 ) ");
                    
                    doc.SendCommand(commandNetload);
                    doc.SendCommand(newCuiPath);
                    doc.SendCommand("\n");
                    doc.SendCommand("(command \"filedia\" 1 ) ");
                    //显示命令行
                    doc.SendCommand("(command \"commandline\" ) ");
                }
            }
            catch (System.Exception ex)
            {
               
            }
        }

        /// <summary>
        /// 卸载程序时的清除操作
        /// </summary>
        public void Terminate()
        {
            
        }

        public static string GetPath()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            return Path.GetDirectoryName(path);
        }              
    }

    public class IniFileHelper
    {
        string strIniFilePath;  // ini配置文件路径  

        // 返回0表示失败，非0为成功  
        [DllImport("kernel32", CharSet = CharSet.Auto)]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        // 返回取得字符串缓冲区的长度  
        [DllImport("kernel32", CharSet = CharSet.Auto)]
        private static extern long GetPrivateProfileString(string section, string key, string strDefault, StringBuilder retVal, int size, string filePath);

        [DllImport("Kernel32.dll", CharSet = CharSet.Auto)]
        public static extern int GetPrivateProfileInt(string section, string key, int nDefault, string filePath);

        /// <summary>  
        /// 无参构造函数  
        /// </summary>  
        /// <returns></returns>  
        public IniFileHelper()
        {
            this.strIniFilePath = IExtension.GetPath() + "\\config.ini";
        }

        /// <summary>
        /// 写INI文件
        /// </summary>
        /// <param name="Section"></param>
        /// <param name="Key"></param>
        /// <param name="Value"></param>
        public void IniWriteValue(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, this.strIniFilePath);
        }
        /// <summary>
        /// 读取INI文件
        /// </summary>
        /// <param name="Section"></param>
        /// <param name="Key"></param>
        /// <returns></returns>
        public string IniReadValue(string Section, string Key, string DefaultValue)
        {
            StringBuilder temp = new StringBuilder(255);
            long i = GetPrivateProfileString(Section, Key, DefaultValue, temp, 255, this.strIniFilePath);
            return temp.ToString();
        }
    }
}

