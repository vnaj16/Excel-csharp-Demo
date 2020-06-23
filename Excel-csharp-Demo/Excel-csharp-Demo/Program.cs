using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;

namespace Excel_csharp_Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            #region CREACION DE PERSONAS A GUARDAR
            Persona obj1 = new Persona(1, "Arthur Valladares", new DateTime(1999, 10, 16));
            Persona obj2 = new Persona(2, "Javier Nole", new DateTime(2000, 10, 14));
            Persona obj3 = new Persona(3, "Graciela Sanchez", new DateTime(2000, 7, 24));
            #endregion


            string PathFile = AppDomain.CurrentDomain.BaseDirectory + "miExcel.xlsx";


            //Depende de DocumentFormat.OpenXml. con la V2.5 y funciona correctamente.
            SLDocument oSLDocument = new SLDocument(); //Documento padre

            //La mejor forma, para aprovechar rendimiento, es esta, usar un DataTable

            DataTable dt = new DataTable();

            //Columns
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Nombre", typeof(string));
            dt.Columns.Add("Fecha Nacimiento", typeof(string));

            //Rows
            dt.Rows.Add(obj1.ID, obj1.Nombre, obj1.Fecha_Nacimiento.ToShortDateString());
            dt.Rows.Add(obj2.ID, obj2.Nombre, obj2.Fecha_Nacimiento.ToShortDateString());
            dt.Rows.Add(obj3.ID, obj3.Nombre, obj3.Fecha_Nacimiento.ToShortDateString());


            //Pongo la tabla al Excel
            //En Excel, se empieza a contar desde 1, no desde 0 como en programacion

            oSLDocument.ImportDataTable(1, 1, dt, true);


            //Guardo en la siguiente ruta

            oSLDocument.SaveAs(PathFile);


            Console.ReadKey();
        }
    }

    public class Persona
    {
        public int ID;
        public string Nombre;
        public DateTime Fecha_Nacimiento;

        public Persona(int id, string name, DateTime FN)
        {
            ID = id;
            Nombre = name;
            Fecha_Nacimiento = FN;
        }
    } 
}
