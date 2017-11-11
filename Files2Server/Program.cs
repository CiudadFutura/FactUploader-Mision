using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinSCP;

namespace Files2Server
{
    class Program
    {


        static void Main(string[] args)
        {

            String certPath = System.IO.Directory.GetCurrentDirectory() + "\\pk.ppk";
            String cfgPath = System.IO.Directory.GetCurrentDirectory() + "\\lastdate.cfg";
            try
            {
                if (!System.IO.File.Exists(certPath)) {
                    Console.WriteLine("Falta la clave privada");
                    return;
                }

                DateTime dt = DateTime.Now;
                String hoy = "";
                if (System.IO.File.Exists(cfgPath))
                {
                    hoy = DateTime.TryParse(System.IO.File.ReadAllText(cfgPath), out dt) ? dt.ToString("_yyMMdd") : "";
                }

                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Scp,
                    HostName = config.getInstance().HostName,
                    UserName = config.getInstance().UserName,
                    PrivateKeyPassphrase = config.getInstance().PrivateKeyPassphrase,
                    SshPrivateKeyPath = System.IO.Directory.GetCurrentDirectory() + "\\pk.ppk",
                    //TlsClientCertificatePath = "pk.ppk",
                    //SshHostKeyFingerprint = "" //ssh-rsa 2048 xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx
                    GiveUpSecurityAndAcceptAnySshHostKey = true,
                };

                using (Session session = new Session())
                {
                    session.FileTransferred += Session_FileTransferred;
                    // Connect
                    session.Open(sessionOptions);
                    Console.WriteLine("Conexion abierta");
                    // Upload files
                    TransferOptions transferOptions = new TransferOptions();
                    transferOptions.TransferMode = TransferMode.Binary;
                    transferOptions.OverwriteMode = OverwriteMode.Overwrite;
                    Console.WriteLine("Subiendo archivos.");
                    TransferOperationResult transferResult;

                    do
                    {
                        transferResult = session.PutFiles(config.getInstance().OrgPath + "*" + hoy + "*.pdf", config.getInstance().DestPath, false, transferOptions);
                        dt = dt.AddDays(1);
                        hoy = dt.ToString("_yyMMdd");
                    } while (dt < DateTime.Now);
                    
                    // Throw on any error
                    transferResult.Check();

                    // Print results
                    foreach (TransferEventArgs transfer in transferResult.Transfers)
                    {
                        Console.WriteLine("Subir {0} terminado", transfer.FileName);
                    }
                }
                System.IO.File.WriteAllText(cfgPath, DateTime.Now.ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);
            }

#if DEBUG
            Console.WriteLine("Fin");
            Console.ReadLine();
#endif
        }

        private static void Session_FileTransferred(object sender, TransferEventArgs e)
        {
            Console.WriteLine("Archivo {0} transferido: {1}", e.FileName, e.Error == null ? "OK" : e.Error.Message);
        }
    }
}
