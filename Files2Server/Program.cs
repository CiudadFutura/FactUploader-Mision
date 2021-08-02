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
		private static readonly string FacturasPath = ConfigurationManager.AppSettings["rutaFacturas"];

		private static void Main()
		{

			String certPath = Directory.GetCurrentDirectory() + "\\pk.ppk";
			String cfgPath = Directory.GetCurrentDirectory() + "\\ultimaExportacionFacturas.cfg";
			String rutaFacturasWeb = ConfigurationManager.AppSettings["rutaFacturasWeb"];

			try
			{
				if (!File.Exists(certPath))
				{
					Console.WriteLine("Falta la clave privada");
					return;
				}

				DateTime dt = DateTime.Now;
				String hoy = "";
				if (File.Exists(cfgPath))
				{
					hoy = DateTime.TryParse(File.ReadAllText(cfgPath), out dt) ? dt.ToString("_yyMMdd") : "";
				}

				//Renombra archivos
				RenombrarArchivosFacturas(hoy, dt);

				var enviar = ConfigurationSettings.AppSettings["enviar"];
				if (enviar != "1") return;

				// Setup session options
				SessionOptions sessionOptions = new SessionOptions
				{
					Protocol = Protocol.Scp,
					HostName = ConfigurationManager.AppSettings["hostName"],
					UserName = ConfigurationManager.AppSettings["userName"],
					PrivateKeyPassphrase = ConfigurationManager.AppSettings["privateKeyPassphrase"],
					SshPrivateKeyPath = certPath,
					//TlsClientCertificatePath = "pk.ppk",
					//SshHostKeyFingerprint = "" //ssh-rsa 2048 xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx
					GiveUpSecurityAndAcceptAnySshHostKey = true,
				};

				using (Session session = new Session())
				{
					session.FileTransferred += Session_FileTransferred;
					// Conectar
					session.Open(sessionOptions);
					Console.WriteLine("Conexion abierta");
					// Actualizar archivos
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

					// Dispara algún error
					transferResult.Check();

					// Mostrar resultados
					foreach (TransferEventArgs transfer in transferResult.Transfers)
					{
						Console.WriteLine("Subir {0} terminado", transfer.FileName);
					}

					Console.WriteLine("Se han enviado {0} archivos de facturas.", transferResult.Transfers.Count);
				}
				File.WriteAllText(cfgPath, DateTime.Now.ToString("yyyy-MM-dd"));
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
			Console.WriteLine("Archivo {0} transferido: {1}", e.FileName, e.Error == null ? "Correcto" : e.Error.Message);
		}

		private static void RenombrarArchivosFacturas(string hoy, DateTime dt)
		{
			var renombrar = ConfigurationSettings.AppSettings["renombrar"];
			if (renombrar != "1") return;

			var rutaArchivosDbf = ConfigurationManager.AppSettings["rutaArchivos"];
			var rutaTablasDbf = ConfigurationManager.AppSettings["rutaTablas"];
			var archivoRemitosDbf = ConfigurationManager.AppSettings["dbfRemitos"];

			//Eliminamos copia previa de MOVART 
			File.Delete(rutaTablasDbf + archivoRemitosDbf);
			//Traemos copia de MOVART porque no podemos acceder al original
			File.Copy(rutaArchivosDbf + archivoRemitosDbf, rutaTablasDbf + archivoRemitosDbf);

			var contador = 0;

			do
			{
				var facturas = Directory.GetFiles(FacturasPath, "FACTLOTR-C-*" + hoy + "*.pdf", SearchOption.TopDirectoryOnly);
				dt = dt.AddDays(1);
				hoy = dt.ToString("_yyMMdd");

				foreach (var factura in facturas)
				{
					try
					{
						var nroMov = factura.Substring(factura.IndexOf("\\FACTLOTR-C-") + 12, 12);
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

						var nuevoNombre = factura.Replace("\\FACTLOTR-C-", "\\FAC_" + nroPedidoWebNum + "_0");
						File.Move(factura, nuevoNombre.Substring(0, nuevoNombre.Length - 30) + ".pdf");

						contador += 1;
					}
					catch (Exception)
					{
						Console.WriteLine("Error renombrando el archivo de factura: {0}", factura);
					}
				}

			} while (dt <= DateTime.Now);

			//Eliminamos copia previa de MOVART 
			File.Delete(rutaTablasDbf + archivoRemitosDbf);

			Console.WriteLine("Se han renombrado {0} archivos de facturas.", contador);
		}
	}
}
