using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace ExportarDuplicados
{
    public class Program
    {

        public static void Main(string[] args)
        {
            // Leer estudiantes desde el Excel
            string filePath = "EstudiantesReinscriptos22_23.xlsx";
            var estudiantes = LeerEstudiantesDesdeExcel(filePath);

            // Insertar los estudiantes en la base de datos
            InsertarEstudiantesEnBaseDeDatos(estudiantes);

            Console.WriteLine("Estudiantes insertados correctamente.");
        }

        public static List<SubjectEnrolmentDTO> LeerEstudiantesDesdeExcel(string filePath)
        {
            var estudiantes = new List<SubjectEnrolmentDTO>();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Supone que los datos están en la primera hoja

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // Asumiendo que los datos comienzan en la fila 2
                {
                    var estudiante = new SubjectEnrolmentDTO
                    {
                        IDSubjectEnrolment = int.Parse(worksheet.Cells[row, 1].Text),
                        AñoInscripcion = int.Parse(worksheet.Cells[row, 2].Text),
                        StudentId = int.Parse(worksheet.Cells[row, 3].Text),
                        NombreEstudiante = worksheet.Cells[row, 4].Text,
                        ApellidoEstudiante = worksheet.Cells[row, 5].Text,
                        Materias = worksheet.Cells[row, 6].Text,
                        Carreras = worksheet.Cells[row, 7].Text,
                        YearInCareer = int.Parse(worksheet.Cells[row, 8].Text)
                    };

                    estudiantes.Add(estudiante);
                }
            }

            return estudiantes;
        }


        public static void InsertarEstudiantesEnBaseDeDatos(List<SubjectEnrolmentDTO> estudiantes)
        {
            using (var db = new Context())
            {
                foreach (var estudianteDto in estudiantes)
                {
                    var materia = db.Subject.FirstOrDefault(s => s.Name == estudianteDto.Materias);
                    if (materia == null)
                    {
                        throw new Exception($"Materia {estudianteDto.Materias} no encontrada en la base de datos.");
                    }

                    var subjectEnrolment = new SubjectEnrolment
                    {
                        StudentId = estudianteDto.StudentId,
                        SubjectId = materia.Id,
                        Year = estudianteDto.AñoInscripcion,
                        Presential = true,
                        Approved = false,
                        CreatedAt = DateTime.Now,
                        LastModificationBy = "Importación desde excel"
                    };

                    // Insertar el registro en la base de datos
                    db.SubjectEnrolment.Add(subjectEnrolment);
                }

                db.SaveChanges();
            }
        }
    }

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

    public class Student
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class Subject
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class SubjectEnrolment
    {
        public int Id { get; set; }
        public int StudentId { get; set; }
        public int SubjectId { get; set; }
        public int Year { get; set; }
        public bool Presential { get; set; }
        public bool Approved { get; set; }
        public DateTime CreatedAt { get; set; }
        public string LastModificationBy { get; set; }

        public Student Student { get; set; }
        public Subject Subject { get; set; }
    }

    public class Context : DbContext
    {
        public DbSet<Student> Student { get; set; }
        public DbSet<Subject> Subject { get; set; }
        public DbSet<SubjectEnrolment> SubjectEnrolment { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("Server=localhost\\SQLEXPRESS;Database=DbGestin;Trusted_Connection=True;TrustServerCertificate=True;");
        }
    }
}

