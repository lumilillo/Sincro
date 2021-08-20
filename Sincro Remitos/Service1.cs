using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Dynamic;
using System.IO;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace Sincro_Remitos
{
    public partial class Service1 : ServiceBase
    {
        #region Parametros

        public static string conString = ConfigurationManager.AppSettings["conString"];
        public static string urlBase = ConfigurationManager.AppSettings["urlBase"];
        public static string api = "/api";
        public static string token = "";
        public static string fecha = "";

        public static int authDni = 1222222222;
        public static string authPassword = "12345678";

        public static string AuthController = "/get_token_user";
        public static string FechaController = "/get_last_update_remito";
        public static string RemitosController = "/created_or_update_remitos";
        public static string TokenController = "/change_token_user";

        Timer timer;
        public static string log = "";
        public static string logPath = AppDomain.CurrentDomain.BaseDirectory + @"/Logs/" + $"{DateTime.Today.ToString("yyyy-MM-dd")}.txt";

        #endregion

        public Service1()
        {
            InitializeComponent();
        }

        #region Timer

        protected override void OnStart(string[] args)
        {
            timer = new Timer(60000);
            timer.Elapsed += new ElapsedEventHandler(timer_tick);
            timer.Enabled = true;
            timer.Start();
        }

        protected override void OnStop()
        {

        }

        async void timer_tick (object sender, ElapsedEventArgs e)
        {
            timer.Enabled = false;
            Log($"***Inicio {DateTime.Now}***");
            Log("\n");

            try
            {
                await Sincronizar();
            }
            catch (Exception ex)
            {
                Log(ex.Message);
                Log("\n");
            }

            Log($"***Fin {DateTime.Now}***");
            Log("\n");
            Log("\n");
            timer.Enabled = true;
            
            this.Dispose();
        }

        #endregion

        #region Funciones
        static async Task Sincronizar()
        {
            Result result = new Result();
            object[] paquete = new object[4];

            //CARGO PAQUETE
            result = await GetToken();
            if (result.isOk) { Log("Token: " + token + "\n"); }
            else { Log(result.error); return; }

            result = await GetFecha(urlBase, api, FechaController, token);
            if (result.isOk) { fecha = result.data.ToString(); }
            else { Log(result.error); return; }

            result = GetClientes();
            if (result.isOk) { paquete[0] = result.data; }
            else { Log(result.error); return; }

            result = GetProductos();
            if (result.isOk) { paquete[1] = result.data; }
            else { Log(result.error); return; }

            result = GetRemitos();
            if (result.isOk) { paquete[2] = result.data; }
            else { Log(result.error); return; }

            result = GetDetalleRemitos();
            if (result.isOk) { paquete[3] = result.data; }
            else { Log(result.error); return; }

            //ENVIO PAQUETE
            result = await Send<object[]>(urlBase, api, RemitosController, paquete, token);
            if (result.isOk) { Log("Sincronización: OK" + "\n"); }
            else { Log(result.error); return; }

            //ACTUALIZO TOKEN
            result = await Send<string>(urlBase, api, TokenController, "", token);
            if (result.isOk) { Log("Token actualizado" + "\n"); }
            else { Log(result.error); return; }
        }

        static async Task<Result> GetToken()
        {
            Result result = new Result();

            try
            {
                Auth auth = new Auth() { dni = authDni, password = authPassword };
                result = await Auth<Auth>(urlBase, api, AuthController, auth);

                if (result.isOk && result.data != null)
                {
                    string apiToken = (string)result.data;

                    string[] split = apiToken.Split(new Char[] { '"' });
                    apiToken = split[1];

                    token = apiToken;
                }
            }
            catch (Exception ex)
            {
                result.error = ex.Message;
            }

            return result;
        }

        static async Task<Result> Auth<T>(
                        string urlBase,
                        string servicePrefix,
                        string controller, T datosUser)
        {
            try
            {
                var json = JsonConvert.SerializeObject(datosUser);
                var data = new StringContent(json, Encoding.UTF8, "application/json");
                var client = new HttpClient();
                client.BaseAddress = new Uri(urlBase);
                var url = $"{servicePrefix}{controller}";
                var response = await client.PostAsync(url, data);
                var result = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    return new Result
                    {
                        isOk = false,
                        data = null,
                        error = result,
                    };
                }

                return new Result
                {
                    isOk = true,
                    data = result,
                    error = null,
                };

            }
            catch (Exception ex)
            {
                return new Result
                {
                    isOk = false,
                    data = null,
                    error = ex.Message,
                };
            }
        }

        static async Task<Result> GetFecha(
                             string urlBase,
                             string servicePrefix,
                             string controller,
                             string apiToken)
        {

            try
            {
                //FECHA DEL SERVIDOR A JSON
                var json = JsonConvert.SerializeObject(apiToken, Formatting.Indented, new IsoDateTimeConverter() { DateTimeFormat = "yyyy-MM-dd HH:mm:ss" });
                var data = new StringContent(json, Encoding.UTF8, "application/json");

                //CLIENTE HTTP
                var client = new HttpClient();
                client.DefaultRequestHeaders.Add("authorization", "Bearer " + apiToken);
                client.DefaultRequestHeaders.Add("Accept", "application/json");
                client.BaseAddress = new Uri(urlBase);

                //URL
                var url = $"{servicePrefix}{controller}";

                //PETICION Y RESPUESTA
                HttpResponseMessage response = await client.PostAsync(url, data);
                var result = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    return new Result
                    {
                        isOk = false,
                        error = result,
                    };
                }

                return new Result
                {
                    isOk = true,
                    error = "Ok",
                    data = result,
                };
            }
            catch (Exception ex)
            {
                return new Result
                {
                    isOk = false,
                    error = ex.Message,
                };
            }

        }

        static async Task<Result> Send<T>(
               string urlBase,
               string servicePrefix,
               string controller,
               T paquete,
               string apiToken)
        {
            try
            {
                //OBJETO A JSON
                var json = JsonConvert.SerializeObject(paquete, Formatting.Indented, new IsoDateTimeConverter() { DateTimeFormat = "yyyy-MM-dd HH:mm:ss" });
                if (paquete is object[]) { Log("Datos enviados: " + "\n" + json + "\n"); }
                //FORMATO DE DATA PARA EVITAR ERRORES DE COMUNICACION
                var data = new StringContent(json, Encoding.UTF8, "application/json");
                //CLIENTE HTTP
                var client = new HttpClient();
                //HEADERS DE LA CONSULTA
                client.DefaultRequestHeaders.Add("authorization", "Bearer " + apiToken);
                client.DefaultRequestHeaders.Add("Accept", "application/json");
                client.BaseAddress = new Uri(urlBase);
                //ARMADO DE LA URL: API/CONTROLADOR
                var url = $"{servicePrefix}{controller}";
                //ENVIO PETICION DE POST URL + DATA (JSON) 
                HttpResponseMessage response = await client.PostAsync(url, data);
                //RESULTADO DE LA CONSULTA
                var result = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    return new Result
                    {
                        isOk = false,
                        error = result,
                    };
                }
                return new Result
                {
                    isOk = true,
                    data = result
                };
            }
            catch (Exception ex)
            {
                return new Result
                {
                    isOk = false,
                    error = ex.Message,
                };
            }
        }

        static Result GetClientes()
        {
            Result result = new Result();

            SqlConnection conexion = new SqlConnection(conString);
            conexion.Open();

            SqlTransaction transaccion = conexion.BeginTransaction();

            try
            {
                SqlCommand comando = new SqlCommand("sp_clientes");
                comando.CommandType = CommandType.StoredProcedure;
                comando.Parameters.Add(new SqlParameter("@fecha_sincronizacion", fecha));
                comando.Connection = conexion;
                comando.Transaction = transaccion;

                SqlDataReader reader = comando.ExecuteReader();

                List<dynamic> objetos = new List<dynamic>();
                while (reader.Read())
                {
                    dynamic objeto = new ExpandoObject();

                    for (int pos = 0; pos < reader.FieldCount; pos++)
                    {
                        AddProperty(objeto, reader.GetName(pos), reader.GetValue(pos));
                    }

                    objetos.Add(objeto);
                }

                reader.Close();
                transaccion.Commit();
                result.isOk = true;
                result.data = objetos;
            }
            catch (Exception ex)
            {
                transaccion.Rollback();
                result.isOk = false;
                result.error = ex.Message;
            }
            finally
            {
                conexion.Close();
            }
            return result;
        }

        static Result GetProductos()
        {
            Result result = new Result();

            SqlConnection conexion = new SqlConnection(conString);
            conexion.Open();

            SqlTransaction transaccion = conexion.BeginTransaction();

            try
            {
                SqlCommand comando = new SqlCommand("sp_productos");
                comando.CommandType = CommandType.StoredProcedure;
                comando.Parameters.Add(new SqlParameter("@fecha_sincronizacion", fecha));
                comando.Connection = conexion;
                comando.Transaction = transaccion;

                SqlDataReader reader = comando.ExecuteReader();

                List<dynamic> objetos = new List<dynamic>();
                while (reader.Read())
                {
                    dynamic objeto = new ExpandoObject();

                    for (int pos = 0; pos < reader.FieldCount; pos++)
                    {
                        AddProperty(objeto, reader.GetName(pos), reader.GetValue(pos));
                    }

                    objetos.Add(objeto);
                }

                reader.Close();
                transaccion.Commit();
                result.isOk = true;
                result.data = objetos;
            }
            catch (Exception ex)
            {
                transaccion.Rollback();
                result.isOk = false;
                result.error = ex.Message;
            }
            finally
            {
                conexion.Close();
            }
            return result;
        }

        static Result GetRemitos()
        {
            Result result = new Result();

            SqlConnection conexion = new SqlConnection(conString);
            conexion.Open();

            SqlTransaction transaccion = conexion.BeginTransaction();

            try
            {
                SqlCommand comando = new SqlCommand("[sp_remitos]");
                comando.CommandType = CommandType.StoredProcedure;
                comando.Parameters.Add(new SqlParameter("@fecha_sincronizacion", fecha));
                comando.Connection = conexion;
                comando.Transaction = transaccion;

                SqlDataReader reader = comando.ExecuteReader();

                List<dynamic> objetos = new List<dynamic>();
                while (reader.Read())
                {
                    dynamic objeto = new ExpandoObject();

                    for (int pos = 0; pos < reader.FieldCount; pos++)
                    {
                        AddProperty(objeto, reader.GetName(pos), reader.GetValue(pos));
                    }

                    objetos.Add(objeto);
                }

                reader.Close();
                transaccion.Commit();
                result.isOk = true;
                result.data = objetos;
            }
            catch (Exception ex)
            {
                transaccion.Rollback();
                result.isOk = false;
                result.error = ex.Message;
            }
            finally
            {
                conexion.Close();
            }
            return result;
        }

        static Result GetDetalleRemitos()
        {
            Result result = new Result();

            SqlConnection conexion = new SqlConnection(conString);
            conexion.Open();

            SqlTransaction transaccion = conexion.BeginTransaction();

            try
            {
                SqlCommand comando = new SqlCommand("[detalle_remitos]");
                comando.CommandType = CommandType.StoredProcedure;
                comando.Parameters.Add(new SqlParameter("@fecha_sincronizacion", fecha));
                comando.Connection = conexion;
                comando.Transaction = transaccion;

                SqlDataReader reader = comando.ExecuteReader();

                List<dynamic> objetos = new List<dynamic>();
                while (reader.Read())
                {
                    dynamic objeto = new ExpandoObject();

                    for (int pos = 0; pos < reader.FieldCount; pos++)
                    {
                        AddProperty(objeto, reader.GetName(pos), reader.GetValue(pos));
                    }

                    objetos.Add(objeto);
                }

                reader.Close();
                transaccion.Commit();
                result.isOk = true;
                result.data = objetos;
            }
            catch (Exception ex)
            {
                transaccion.Rollback();
                result.isOk = false;
                result.error = ex.Message;
            }
            finally
            {
                conexion.Close();
            }
            return result;
        }

        public static void AddProperty(ExpandoObject expando, string propertyName, object propertyValue)
        {
            var expandoDict = expando as IDictionary<string, object>;

            if (expandoDict.ContainsKey(propertyName))
            {
                expandoDict[propertyName] = propertyValue;
            }
            else
            {
                expandoDict.Add(propertyName, propertyValue);
            }
        }

        public static void Log(string linea)
        {
            DirectoryInfo dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "Logs");

            if (!dir.Exists)
            {
                dir.Create();
            }

            File.AppendAllText(logPath, linea);
        }

        #endregion

    }
}
