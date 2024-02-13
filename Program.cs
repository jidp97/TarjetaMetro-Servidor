using OfficeOpenXml;
using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Timers;

namespace MetroCardSimulator
{
    class Program
    {
        static Dictionary<int, MovimientoTarjeta> movimientosTarjeta = new Dictionary<int, MovimientoTarjeta>();
        static System.Timers.Timer timer;

        static async Task Main(string[] args)
        {
            // Establece el contexto de licencia de ExcelPackage a no comercial
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Crea un HttpListener para escuchar las solicitudes entrantes
            var httpListener = new HttpListener();

            // Agrega la URL de prefijo para la escucha
            httpListener.Prefixes.Add("http://localhost:9090/");

            // Inicia el listener
            httpListener.Start();

            // Imprime un mensaje en la consola indicando que el servidor WebSocket está iniciado
            Console.WriteLine("Servidor WebSocket iniciado. Esperando conexiones...");

            // Crea un temporizador que se activará cada 2 minutos
            timer = new System.Timers.Timer(120000); // 120000 milisegundos = 2 minutos

            // Asigna el método OnTimedEvent para manejar los eventos del temporizador
            timer.Elapsed += OnTimedEvent;

            // Configura el temporizador para que se reinicie automáticamente y lo habilita
            timer.AutoReset = true;
            timer.Enabled = true;

            // Entra en un bucle infinito para manejar las solicitudes entrantes
            while (true)
            {
                // Espera una solicitud HTTP entrante
                var context = await httpListener.GetContextAsync();

                // Verifica si la solicitud es una solicitud WebSocket
                if (context.Request.IsWebSocketRequest)
                {
                    // Acepta la conexión WebSocket
                    var webSocketContext = await context.AcceptWebSocketAsync(null);

                    // Maneja la conexión WebSocket de forma asincrónica
                    await HandleWebSocketAsync(webSocketContext.WebSocket);
                }
                else
                {
                    // Rechaza las solicitudes que no sean WebSocket
                    context.Response.StatusCode = 400;
                    context.Response.Close();
                }
            }


            
        }

        static async Task HandleWebSocketAsync(WebSocket webSocket)
        {
            // Saldo actual de la tarjeta
            double saldo = 0.0;

            // Piso destino en el contexto de la tarjeta del metro
            int estacionDestino = 1;

            // Tipo de movimiento de la tarjeta (Recarga o Consumo)
            MovimientoTipo tipoMovimiento = MovimientoTipo.Recarga;

            // Buffer para recibir mensajes WebSocket
            byte[] buffer = new byte[1024];

            // Recibe el primer mensaje del cliente
            WebSocketReceiveResult result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);

            // Bucle principal que maneja la comunicación WebSocket
            while (!result.CloseStatus.HasValue)
            {
                // Decodifica el mensaje recibido como una cadena UTF-8
                string message = Encoding.UTF8.GetString(buffer, 0, result.Count);
                Console.WriteLine($"Mensaje recibido: {message}");

                // Maneja el tipo de movimiento según el protocolo definido en los mensajes
                if (message.StartsWith("Recarga"))
                {
                    tipoMovimiento = MovimientoTipo.Recarga;
                }
                else if (message.StartsWith("Consumo"))
                {
                    tipoMovimiento = MovimientoTipo.Consumo;
                }

                // Intenta convertir el mensaje en un número (que representa el monto de recarga o consumo)
                if (double.TryParse(message.Substring(7), out double monto))
                {
                    // Realiza la acción correspondiente según el tipo de movimiento
                    if (tipoMovimiento == MovimientoTipo.Recarga)
                    {
                        saldo += monto;
                    }
                    else if (tipoMovimiento == MovimientoTipo.Consumo)
                    {
                        if (saldo >= monto)
                        {
                            saldo -= monto;
                            estacionDestino = (estacionDestino % 10) + 1; // Simula movimiento a una estación aleatoria
                        }
                        else
                        {
                            // Si no hay suficiente saldo, envía un mensaje de error al cliente
                            string errorMessage = "\u001b[31mSaldo insuficiente para el viaje.\u001b[0m";
                            byte[] errorBuffer = Encoding.UTF8.GetBytes(errorMessage);
                            await webSocket.SendAsync(new ArraySegment<byte>(errorBuffer), WebSocketMessageType.Text, true, CancellationToken.None);
                        }
                    }

                    // Registra el movimiento en la tarjeta
                    MovimientoTarjeta movimiento = new MovimientoTarjeta
                    {
                        Tipo = tipoMovimiento,
                        Monto = monto,
                        SaldoRestante = saldo
                    };

                    movimientosTarjeta.Add(movimientosTarjeta.Count + 1, movimiento);

                    // Envía una actualización de estado al cliente
                    string response = $"\u001b[36mSaldo actual: {saldo}. Destino: Estación {estacionDestino}.\u001b[0m";
                    byte[] responseBuffer = Encoding.UTF8.GetBytes(response);
                    await webSocket.SendAsync(new ArraySegment<byte>(responseBuffer), WebSocketMessageType.Text, true, CancellationToken.None);
                }

                // Recibe el siguiente mensaje del cliente
                buffer = new byte[1024];
                result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);
            }

            // Cierra la conexión WebSocket
            await webSocket.CloseAsync(result.CloseStatus.Value, result.CloseStatusDescription, CancellationToken.None);
        }

        private static void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            // Llama a la función para guardar los datos en Excel.
            GuardarMovimientosEnExcel();
        }

        private static void GuardarMovimientosEnExcel()
        {
            string fileName = "movimientos_tarjeta.xlsx"; // Nombre del archivo
            string directory = AppDomain.CurrentDomain.BaseDirectory; // Obtiene la ruta del directorio actual

            string filePath = Path.Combine(directory, fileName); // Combina la ruta del directorio con el nombre del archivo

            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                file.Delete(); // Elimina el archivo existente para escribir uno nuevo
            }

            // Crea un nuevo archivo de Excel y guarda la información de movimientos en él.
            using (ExcelPackage package = new ExcelPackage(file))
            {
                // Crea una nueva hoja de trabajo llamada "Movimientos".
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Movimientos");

                // Agrega encabezados a la hoja de trabajo.
                worksheet.Cells[1, 1].Value = "Tipo";
                worksheet.Cells[1, 2].Value = "Monto";
                worksheet.Cells[1, 3].Value = "Saldo Restante";

                // Recorre los datos de movimientos y los agrega a la hoja de trabajo.
                int row = 2;  // Comienza a escribir datos en la fila 2 (después de los encabezados)
                foreach (var movimiento in movimientosTarjeta.Values)
                {
                    worksheet.Cells[row, 1].Value = movimiento.Tipo.ToString();
                    worksheet.Cells[row, 2].Value = movimiento.Monto;
                    worksheet.Cells[row, 3].Value = movimiento.SaldoRestante;
                    row++;  // Avanza a la siguiente fila
                }

                // Guarda los cambios en el archivo de Excel.
                package.Save();
            }

            // Imprime un mensaje en la consola indicando la ubicación y la fecha/hora de guardado del archivo.
            Console.WriteLine($"Datos guardados en {filePath} - {DateTime.Now}");
        }
    }

    // Enumeración para representar el tipo de movimiento de la tarjeta (Recarga o Consumo)
    enum MovimientoTipo
    {
        Recarga,
        Consumo
    }

    // Clase que representa información sobre los movimientos de una tarjeta del metro.
    class MovimientoTarjeta
    {
        // Tipo de movimiento (Recarga o Consumo)
        public MovimientoTipo Tipo { get; set; }

        // Monto de la recarga o consumo
        public double Monto { get; set; }

        // Saldo restante en la tarjeta después del movimiento
        public double SaldoRestante { get; set; }
    }
}
