using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExportarDuplicados
{
    public class Program
    {
        private static string connectionString2022 = "Server=localhost\\SQLEXPRESS;Database=dbGestin22;Trusted_Connection=True;";
        private static string connectionString2023 = "Server=localhost\\SQLEXPRESS;Database=dbGestin23;Trusted_Connection=True;";
        private static string connectionString2024 = "Server=localhost\\SQLEXPRESS;Database=DbGestin;Trusted_Connection=True;";

        public static void Main(string[] args)
        {
            // Obtener estudiantes de todas las bases de datos
            var estudiantesReinscriptos = ObtenerEstudiantes();

            // Exportar los estudiantes reinscriptos a un archivo Excel
            ExportarAExcel(estudiantesReinscriptos, "EstudiantesReinscriptos.xlsx");

            Console.WriteLine("Archivo Excel exportado correctamente.");
        }

        // Obtener estudiantes de las bases de datos 2022, 2023 y 2024
        public static List<SubjectEnrolmentDTO> ObtenerEstudiantes()
        {
            var estudiantes = new List<SubjectEnrolmentDTO>();

            estudiantes.AddRange(ObtenerEstudiantesPorYear(2024, connectionString2024));

            var estudiantes2023 = ObtenerEstudiantesPorYear(2023, connectionString2023);

            var estudiantes2022 = ObtenerEstudiantesPorYear(2022, connectionString2022);

            // Obtener estudiantes reinscriptos
            var estudiantesReinscritos = estudiantes2022
                .Where(e2022 => estudiantes2023.Any(e2023 => e2023.StudentId == e2022.StudentId && e2023.Materias == e2022.Materias))
                .Distinct() 
                .ToList();

            return estudiantesReinscritos;
        }


        // Obtener estudiantes de la base de datos por año
        private static List<SubjectEnrolmentDTO> ObtenerEstudiantesPorYear(int año, string connectionString)
        {
            var estudiantes = new List<SubjectEnrolmentDTO>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = @"
                    SELECT 
                        se.Id AS IDSubjectEnrolment,
                        se.Year AS AñoInscripcion,
                        st.Id AS StudentId,
                        u.Name AS NombreEstudiante,
                        u.LastName AS ApellidoEstudiante,
                        s.Name AS Materias,
                        c.Name AS Carreras,
                        s.YearInCareer
                    FROM SubjectEnrolment se
                        INNER JOIN Subject s ON se.SubjectId = s.Id
                        INNER JOIN Student st ON se.StudentId = st.Id
                        INNER JOIN [User] u ON st.UserId = u.Id
                        INNER JOIN Career c ON s.CareerId = c.Id
                    WHERE se.Year = @Year 
                        AND s.DeletedAt IS NULL 
                    ORDER BY st.Id ASC";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Year", año);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var estudiante = new SubjectEnrolmentDTO
                            {
                                IDSubjectEnrolment = reader.GetInt32(0),
                                AñoInscripcion = reader.GetInt32(1),
                                StudentId = reader.GetInt32(2),
                                NombreEstudiante = reader.GetString(3),
                                ApellidoEstudiante = reader.GetString(4),
                                Materias = reader.GetString(5),
                                Carreras = reader.GetString(6),
                                YearInCareer = reader.GetInt32(7)
                            };

                            estudiantes.Add(estudiante);
                        }
                    }
                }
            }

            return estudiantes;
        }

        // Exportar los estudiantes reinscriptos a Excel
        public static void ExportarAExcel(List<SubjectEnrolmentDTO> estudiantesReinscriptos, string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheetReinscriptos = package.Workbook.Worksheets.Add("Estudiantes Reinscriptos");

                worksheetReinscriptos.Cells[1, 1].Value = "ID Inscripción";
                worksheetReinscriptos.Cells[1, 2].Value = "Año Inscripción";
                worksheetReinscriptos.Cells[1, 3].Value = "ID Estudiante";
                worksheetReinscriptos.Cells[1, 4].Value = "Nombre";
                worksheetReinscriptos.Cells[1, 5].Value = "Apellido";
                worksheetReinscriptos.Cells[1, 6].Value = "Materia";
                worksheetReinscriptos.Cells[1, 7].Value = "Carrera";
                worksheetReinscriptos.Cells[1, 8].Value = "Año Carrera";

                int row = 2;
                foreach (var estudiante in estudiantesReinscriptos)
                {
                    worksheetReinscriptos.Cells[row, 1].Value = estudiante.IDSubjectEnrolment;
                    worksheetReinscriptos.Cells[row, 2].Value = estudiante.AñoInscripcion;
                    worksheetReinscriptos.Cells[row, 3].Value = estudiante.StudentId;
                    worksheetReinscriptos.Cells[row, 4].Value = estudiante.NombreEstudiante;
                    worksheetReinscriptos.Cells[row, 5].Value = estudiante.ApellidoEstudiante;
                    worksheetReinscriptos.Cells[row, 6].Value = estudiante.Materias;
                    worksheetReinscriptos.Cells[row, 7].Value = estudiante.Carreras;
                    worksheetReinscriptos.Cells[row, 8].Value = estudiante.YearInCareer;
                    row++;
                }

                worksheetReinscriptos.Cells.AutoFitColumns();

                // Guardar el archivo Excel
                package.SaveAs(new FileInfo(filePath));
            }
        }
    }

    // DTO para representar los datos de inscripción de los estudiantes
    public class SubjectEnrolmentDTO
    {
        public int IDSubjectEnrolment { get; set; }
        public int AñoInscripcion { get; set; }
        public int StudentId { get; set; }
        public string NombreEstudiante { get; set; }
        public string ApellidoEstudiante { get; set; }
        public string Materias { get; set; }
        public string Carreras { get; set; }
        public int YearInCareer { get; set; }
    }
}
