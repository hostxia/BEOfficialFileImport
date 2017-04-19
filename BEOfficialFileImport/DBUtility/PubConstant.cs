using System.Configuration;

namespace BEOfficialFileImport.DBUtility
{

    public class PubConstant
    {
        /// <summary>
        /// 获取连接字符串
        /// </summary>
        public static string ConnectionStringCPC
        {
            get
            {
                return ConfigurationManager.AppSettings["ConnectionStringCPC"];
            }
        }

        /// <summary>
        /// 获取连接字符串
        /// </summary>
        public static string ConnectionStringPC
        {
            get
            {
                return ConfigurationManager.AppSettings["ConnectionStringPC"];
            }
        }

        /// <summary>
        /// 得到web.config里配置项的数据库连接字符串。
        /// </summary>
        /// <param name="configName"></param>
        /// <returns></returns>
        public static string GetConnectionString(string configName)
        {
            return ConfigurationManager.AppSettings[configName];
        }


    }
}
