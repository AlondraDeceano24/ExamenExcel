using PruebaExcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;

public class Program
{
    static void Main(string[] args)
    {

        Practica practica = new Practica();
        
       
        (List<Registro> registrosVerdaderos, int completados) = Practica.LeerExcel();

      
        List<Registro> registros = registrosVerdaderos
            .GroupBy(r => new { r.Fecha, r.Argentina, r.Brasil, r.Chile, r.Peru })
            .Select(g => g.First())
            .ToList();

   

        int registrosDuplicados = registrosVerdaderos.Count - registros.Count;


        var agrupados = registros
        .GroupBy(row => new
        {
            Año = row.Fecha.Year,
            Semana = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(row.Fecha, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)


        })
        .Select(group => new
        {
                Fecha = group.Min(r => r.Fecha),
                VentasArgentina = group.Sum(r => r.Argentina),
                VentasBrasil = group.Sum(r => r.Brasil),
                VentasChile = group.Sum(r => r.Chile),
                VentasPeru = group.Sum(r => r.Peru),
                VentasTotal = group.Sum(r => r.Argentina + r.Brasil + r.Peru + r.Chile)



        })
      .OrderBy(group => group.Fecha)
      .ToList();


    
        foreach (var registro in agrupados)
        {
            Console.WriteLine($"Fecha: {registro.Fecha}, Brasil: {registro.VentasBrasil}, Chile: {registro.VentasChile}, Peru: {registro.VentasPeru}, Argentina: {registro.VentasArgentina}");
        }

        Console.WriteLine("Los registros eliminados fueron: " + registrosDuplicados);
        Console.WriteLine("Los registros completados fueron con 0: " + completados);
        Console.ReadKey();
    }
}

public class Practica
{
 
    public static ( List<Registro>,int) LeerExcel()

    {
        int contador = 0;
        string filePath = @"C:\Users\digis\OneDrive\Documento\DocumentoExcel.xlsx";
        string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";";

        List<Registro> registros = new List<Registro>();

        try
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("SELECT * FROM [Sheet1$]", connection);
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

             
                foreach (DataRow row in dataTable.Rows)
                {
                    Registro registro = new Registro();


                    if (DateTime.TryParse(row[0]?.ToString(), out DateTime fecha))
                    {
                        if(string.IsNullOrEmpty(row[1]?.ToString()) || string.IsNullOrEmpty(row[2]?.ToString()) || string.IsNullOrEmpty(row[3]?.ToString()) || string.IsNullOrEmpty(row[4]?.ToString())) 
                        {
                            contador++;
                        }

                        
                        registro.Fecha = fecha;
                        registro.Argentina = string.IsNullOrEmpty(row[1]?.ToString()) ? 0 : Convert.ToInt32(row[1]);
                        registro.Brasil = string.IsNullOrEmpty(row[2]?.ToString()) ? 0 : Convert.ToInt32(row[2]);
                        registro.Peru = string.IsNullOrEmpty(row[3]?.ToString()) ? 0 : Convert.ToInt32(row[3]);
                        registro.Chile = string.IsNullOrEmpty(row[4]?.ToString()) ? 0 : Convert.ToInt32(row[4]);

                        registros.Add(registro);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ocurrió un error al leer el archivo: " + ex.Message);
        }

        return (registros, contador);
    }
}


