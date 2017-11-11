using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Soap;

namespace Files2Server
{
    class config
    {
        [Serializable]
        public class storage
        {
            public string PrivateKeyPassphrase = "pass";
            public string HostName = "domain";
            public string UserName = "usr";
            public string OrgPath = "E:\\hola";
            public string DestPath = "/home/hola";
        }

        static config INSTANCE;

        storage st;
        String cfgPath = System.IO.Directory.GetCurrentDirectory() + "\\cfg.xml";

        private config() {
            
            IFormatter formatter = new SoapFormatter();
            Stream stream;
            if (System.IO.File.Exists(cfgPath))
            {
                stream = new FileStream(cfgPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                st = (storage)formatter.Deserialize(stream);
            }
            else {
                st = new storage();
                stream = new FileStream(cfgPath, FileMode.Create, FileAccess.Write, FileShare.None);
                formatter.Serialize(stream, st);
            }
            stream.Close();

        }

        public static config getInstance() {
            return INSTANCE != null ? INSTANCE : INSTANCE = new config();
        }

        public string PrivateKeyPassphrase
        {
            get
            {
                return st.PrivateKeyPassphrase;
            }
        }
        public string HostName {
            get
            {
                return st.HostName;
            }
        }
        public string UserName {
            get
            {
                return st.UserName;
            }
        }

        public string OrgPath
        {
            get
            {
                return st.OrgPath;
            }
        }

        public string DestPath
        {
            get
            {
                return st.DestPath;
            }
        }
    }
}
