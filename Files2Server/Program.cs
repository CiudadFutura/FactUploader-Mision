using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using WinSCP;

namespace Files2Server
{
    internal class Program
    {
        private static readonly string FacturasPath = ConfigurationManager.AppSettings["rutaFacturas"].ToString(); 

        private static void Main(string[] args)
        {

            String certPath = System.IO.Directory.GetCurrentDirectory() + "\\pk.ppk";
            String cfgPath = System.IO.Directory.GetCurrentDirectory() + "\\ultimaExportacionFacturas.cfg";
            String rutaFacturasWeb = ConfigurationManager.AppSettings["rutaFacturasWeb"].ToString();

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

                //Renombra archivos
                RenombrarArchivosFacturas(hoy, dt);

                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Scp,
                    HostName = ConfigurationManager.AppSettings["hostName"].ToString(),
                    UserName = ConfigurationManager.AppSettings["userName"].ToString(),
                    PrivateKeyPassphrase = ConfigurationManager.AppSettings["privateKeyPassphrase"].ToString(),
                    SshPrivateKeyPath = certPath,
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
                    TransferOptions transferOptions = new TransferOptions
                    {
                        TransferMode = TransferMode.Binary,
                        OverwriteMode = OverwriteMode.Overwrite,
                    };
                    Console.WriteLine("Subiendo archivos.");
                    TransferOperationResult transferResult;

                    do
                    {
                        transferResult = session.PutFiles(FacturasPath + "FAC_*" + hoy + "*.pdf", rutaFacturasWeb, false, transferOptions);
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
                File.WriteAllText(cfgPath, DateTime.Now.ToString());
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

        private static void RenombrarArchivosFacturas(string hoy, DateTime dt)
        {

            var rutaArchivosDbf = ConfigurationManager.AppSettings["rutaArchivos"].ToString();
            var rutaTablasDbf = ConfigurationManager.AppSettings["rutaTablas"].ToString();
            var archivoRemitosDbf = ConfigurationManager.AppSettings["dbfRemitos"].ToString();

            File.Delete(rutaTablasDbf + archivoRemitosDbf);
            File.Copy(rutaArchivosDbf + archivoRemitosDbf, rutaTablasDbf + archivoRemitosDbf);

            do
            {
                var facturas = Directory.GetFiles(FacturasPath, "FAC0*" + hoy + "*.pdf", SearchOption.TopDirectoryOnly);
                dt = dt.AddDays(1);
                hoy = dt.ToString("_yyMMdd");

                foreach (var factura in facturas)
                {
                    try
                    {
                        var nroMov = factura.Substring(factura.IndexOf("\\FAC0") + 4, 12);
                        Console.WriteLine("Renombrando el archivo de factura: " + factura);
                        //SELECCIONAR nroMovRemito a partir de nroMovFactura
                        var sConArchivos = ConfigurationManager.AppSettings["stringConexion"].Replace("{0}", rutaTablasDbf);
                        var sqlRemitoPorNroMovFactura = string.Format(ConfigurationManager.AppSettings["remitoPorNroMovFactura"],
                            archivoRemitosDbf, nroMov);
                        var tabla = new DataTable();
                        //Console.WriteLine(sqlRemitoPorNroMovFactura);
                        var conexion = new OleDbConnection(sConArchivos);
                        conexion.Open();

                        var comando = new OleDbCommand
                        {
                            Connection = conexion,
                            CommandText = sqlRemitoPorNroMovFactura,
                            CommandType = CommandType.Text
                        };

                        var da = new OleDbDataAdapter(comando);
                        da.Fill(tabla);
                        var nroRemito = tabla.AsEnumerable().Select(r => r.Field<decimal>("NROMOVI")).FirstOrDefault();
                        conexion.Close();
                        //Console.WriteLine(nroRemito);
                        if (nroRemito == 0) continue;

                        //SELECCIONAR nroPedidoWeb a partir de nroMovRemito
                        var sConTablas = ConfigurationManager.AppSettings["stringConexion"].Replace("{0}", rutaTablasDbf);
                        var archivoDbf = ConfigurationManager.AppSettings["dbfPedidos"];
                        var sqlPedidoPorNroMov = string.Format(ConfigurationManager.AppSettings["pedidoPorNroMov"], archivoDbf,
                            nroRemito);
                        tabla = new DataTable();
                        //Console.WriteLine(sqlPedidoPorNroMov);
                        conexion = new OleDbConnection(sConTablas);
                        conexion.Open();

                        comando = new OleDbCommand
                        {
                            Connection = conexion,
                            CommandText = sqlPedidoPorNroMov,
                            CommandType = CommandType.Text
                        };
                        da = new OleDbDataAdapter(comando);
                        da.Fill(tabla);
                        var nroPedidoWeb = tabla.AsEnumerable().Select(r => r.Field<string>("NROPEDIDO").TrimEnd()).FirstOrDefault();

                        conexion.Close();
                        //Console.WriteLine(nroPedidoWeb);
                        long nroPedidoWebNum;
                        if (!long.TryParse(nroPedidoWeb, out nroPedidoWebNum)) continue;

                        File.Move(factura, factura.Replace("\\FAC0", "\\FAC_" + nroPedidoWebNum + "_0"));
                    }
                    catch (Exception)
                    {
                        //Console.WriteLine("Error renombrando el archivo de factura: " + factura);
                    }
                }

            } while (dt <= DateTime.Now);

            File.Delete(rutaTablasDbf + archivoRemitosDbf);
        }
    }
}
