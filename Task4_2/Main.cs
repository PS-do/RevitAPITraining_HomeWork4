using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using File = System.IO.File;

namespace Task4_2
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

// Выведите в файл Excel следующие значения всех труб:
// имя типа трубы,
// наружный диаметр трубы,
// внутренний диаметр трубы,
// длина трубы.

            string pipesInfo = string.Empty;
            List<Pipe> pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();

            var saveFileDialog = new SaveFileDialog
            {
                OverwritePrompt = true,//если файл существует, выдавать запрос на его перезапись
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),//начальная папка, скоторой будет начинаться диалог
                Filter = "All files(*.*)|*.*",//фильтр отображаемых файлов по расширению
                FileName = "pipesInfo.хlsx",//Имя файла по умолчанию
                DefaultExt = ".хlsx"//Расширение по умолчани
            };
            string excelPatch = string.Empty;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                excelPatch = saveFileDialog.FileName;
            }
            if (string.IsNullOrEmpty(excelPatch))
                return Result.Cancelled;



            using (FileStream fs = new FileStream(excelPatch, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workBook = new XSSFWorkbook();
                ISheet sheet = workBook.CreateSheet("Лист1");
                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.PipeType.Name);
                    sheet.SetCellValue(rowIndex, columnIndex: 1,
                        UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble(), UnitTypeId.Meters));// наружный диаметр трубы,
                    sheet.SetCellValue(rowIndex, columnIndex: 2,
                        UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble(), UnitTypeId.Meters));// внутренний диаметр трубы,
                    sheet.SetCellValue(rowIndex, columnIndex: 3,
                        UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), UnitTypeId.Meters));// длина трубы.
                    rowIndex++;
                }
                workBook.Write(fs);
                workBook.Close();
            }

            System.Diagnostics.Process.Start(excelPatch);


            return Result.Succeeded;
        }
    }
}
