using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.IO;
using System.Xml;
using System.Reflection;
using System.Diagnostics;
namespace acad
{

    /// <summary>    
    /// 支持DBSettings和ConnectionStrings节点 ,如需其他节点请自行扩展
    /// </summary>    
    public class APIConfHelper
    {
        //实现 DBSettings 节点的读取          
        public static Hashtable DBSettings
        {
            get
            {
                return GetNameAndValue("DBSettings", "key", "value");
            }
        }
        /// <summary>       
        /// 实现对DBSettings节点对应key的value设置        
        /// </summary>        
        /// <param name="keyNameValue"></param>        
        /// <param name="value"></param>     
        public static void SetDBSettings(string keyNameValue, string value)
        {

            SetNameAndValue("DBSettings", "key", keyNameValue, value);
        }     
       
        /// <summary>
        /// 实现对相应节点的对应节点对应key或者name对应value值的设置        
        /// </summary>        
        /// <param name="sectionTag">对应节点</param>        
        /// <param name="KeyOrName">对应key或者name</param>        
        /// <param name="keyNameValue">key或者name的值</param>        
        /// <param name="valueOrConnectionString">对应key或者name 的value值</param>        
        private static void SetNameAndValue(string sectionTag, string KeyOrName, string keyNameValue, string valueOrConnectionString)
        {
            string assemblyPath = Assembly.GetExecutingAssembly().GetName().CodeBase;
            //获取运行项目当然DLL的路径            
            assemblyPath = assemblyPath.Remove(0, 8);
            //去除路径前缀            
            string configUrl = assemblyPath + ".config";
            //添加.config后缀，得到配置文件路径           
            try
            {
                XmlDocument cfgDoc = new XmlDocument();
                cfgDoc.Load(configUrl);
                XmlNodeList nodes = cfgDoc.GetElementsByTagName(sectionTag);
                foreach (XmlNode node in nodes)
                {
                    foreach (XmlNode childNode in node.ChildNodes)
                    {
                        XmlAttributeCollection attributes = childNode.Attributes;
                        if (attributes != null)
                        {
                            if (attributes.GetNamedItem(KeyOrName).InnerText == keyNameValue)
                            {
                                attributes.GetNamedItem("value").InnerText = valueOrConnectionString;
                            }
                        }
                    }
                }
                cfgDoc.Save(configUrl);
            }
            catch (FileNotFoundException es)
            {
                throw es;
            }
        }
        /// <summary>
        /// 根据节点名，子节点名，获取指定值           
        /// </summary>           
        /// <param name="sectionTag">对应节点</param>           
        /// <param name="KeyOrName">key或者name</param>           
        /// <param name="valueOrConnectionString">key或者name的值</param>           
        /// <returns>key或者name对应value值</returns>           
        private static Hashtable GetNameAndValue(string sectionTag, string KeyOrName, string valueOrConnectionString)
        {
            Hashtable settings = new Hashtable(5);//初始化Hashtable             
            string assemblyPath = Assembly.GetExecutingAssembly().GetName().CodeBase;//获取运行项目当前DLL的路径            
            assemblyPath = assemblyPath.Remove(0, 8); //去除前缀            
            string configUrl = assemblyPath + ".config"; //添加 .config 后缀，得到配置文件路径                
            XmlDocument cfgDoc = new XmlDocument();
            FileStream fs = null;
            try
            {
                fs = new FileStream(configUrl, FileMode.Open, FileAccess.Read);
                cfgDoc.Load(new XmlTextReader(fs));
                XmlNodeList nodes = cfgDoc.GetElementsByTagName(sectionTag);
                foreach (XmlNode node in nodes)
                {
                    foreach (XmlNode childNode in node.ChildNodes)
                    {
                        XmlAttributeCollection attributes = childNode.Attributes;
                        if (attributes != null)
                        {                            //为null不添加                            
                            settings.Add(attributes[KeyOrName].Value, attributes[valueOrConnectionString].Value);
                        }
                    }
                }
            }
            catch (FileNotFoundException es)
            {
                throw es;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
            return settings;
        }
    }
}