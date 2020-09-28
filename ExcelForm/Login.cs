using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
//using System.ServiceModel;
//using System.ServiceModel.Description;
using System.Text;
using System.Windows.Forms;
using System.ServiceModel.Web;
using System.Net;
using System.Reflection;
using System.CodeDom.Compiler;
using Microsoft.CSharp;
using System.CodeDom;
using System.IO;
using System.Web.Services.Description;


namespace ExportTableToExcel
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            

            //WebServiceHost webServiceHost = new WebServiceHost(typeof(ServiceReference1.JoydriveSoapClient), new Uri[] { new Uri("http://acd789.joydrive.cn/joydrive.asmx") });
            //WebHttpBinding webBinding = new WebHttpBinding();
            //webServiceHost.AddServiceEndpoint(typeof(ServiceReference1.JoydriveSoap), webBinding, "");
        }
        public static bool isSuccess = false;
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {//登陆
            //服务地址，该地址可以放到程序的配置文件中，这样即使服务地址改变了，也无须重新编译程序。
            string url = "http://acd789.joydrive.cn/joydrive.asmx";
            //客户端代理服务命名空间，可以设置成需要的值。
            string ns = string.Format("ProxyServiceReference");

            //获取WSDL
            WebClient wc = new WebClient();
            Stream stream = wc.OpenRead(url + "?WSDL");
            ServiceDescription sd = ServiceDescription.Read(stream);//服务的描述信息都可以通过ServiceDescription获取
            string classname = sd.Services[0].Name;

            ServiceDescriptionImporter sdi = new ServiceDescriptionImporter();
            sdi.AddServiceDescription(sd, "", "");
            CodeNamespace cn = new CodeNamespace(ns);

            //生成客户端代理类代码
            CodeCompileUnit ccu = new CodeCompileUnit();
            ccu.Namespaces.Add(cn);
            sdi.Import(cn, ccu);
            CSharpCodeProvider csc = new CSharpCodeProvider();

            //设定编译参数
            CompilerParameters cplist = new CompilerParameters();
            cplist.GenerateExecutable = false;
            cplist.GenerateInMemory = true;
            cplist.ReferencedAssemblies.Add("System.dll");
            cplist.ReferencedAssemblies.Add("System.XML.dll");
            cplist.ReferencedAssemblies.Add("System.Web.Services.dll");
            cplist.ReferencedAssemblies.Add("System.Data.dll");

            //编译代理类
            CompilerResults cr = csc.CompileAssemblyFromDom(cplist, ccu);
            if (cr.Errors.HasErrors == true)
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                foreach (System.CodeDom.Compiler.CompilerError ce in cr.Errors)
                {
                    sb.Append(ce.ToString());
                    sb.Append(System.Environment.NewLine);
                }
                throw new Exception(sb.ToString());
            }

            //生成代理实例，并调用方法
            Assembly assembly = cr.CompiledAssembly;
            Type t = assembly.GetType(ns + "." + classname, true, true);
            object obj = Activator.CreateInstance(t);
            MethodInfo add = t.GetMethod("Login");
            string a = "1", b = "1";//Add方法的参数
            object[] addParams = new object[] { UserName.Text, UserPasswd.Text };
            object addReturn = add.Invoke(obj, addParams);
            if(addReturn == null || !((bool)addReturn))
            {
                if(addReturn == null) MessageBox.Show("登陆异常！");
                isSuccess = false;
                MessageBox.Show("登陆失败！");
            }
            else
            {
                isSuccess = true;
                this.Close();
            }
            //ExportTableToExcel.ServiceReference1.JoydriveSoapClient joydriveSoapClient = new ExportTableToExcel.ServiceReference1.JoydriveSoapClient("JoydriveSoap");
            //isSuccess = joydriveSoapClient.Login(UserName.Text, UserName.Text);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {//注册
            Process pro = new Process();
            pro.StartInfo.FileName = "iexplore.exe";
            pro.StartInfo.Arguments = "http://cloud.joydrive.cn/xtzxk/reguser.aspx";
            pro.Start();
        }
    }
}
