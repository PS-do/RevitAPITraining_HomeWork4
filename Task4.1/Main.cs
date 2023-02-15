using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Task4_1
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
//Выведите в текстовый файл с разделением данных(разделитель может быть любой) следующие
//значения всех стен: имя типа стены, объём стены.

            string wallInfo = string.Empty;
            List<Wall> walls = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            foreach (Wall wall in walls)
            {
                string typeName = wall.WallType.Name;
                string volume = /*wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsValueString();*/
                            UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble(), UnitTypeId.CubicMeters).ToString();
                wallInfo += $"{typeName}\t{volume}{Environment.NewLine}";
            }

            var saveFileDialog = new SaveFileDialog
            {
                OverwritePrompt = true,//если файл существует, выдавать запрос на его перезапись
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),//начальная папка, скоторой будет начинаться диалог
                Filter = "All files(*.*)|*.*",//фильтр отображаемых файлов по расширению
                FileName = "wallInfo.csv",//Имя файла по умолчанию
                DefaultExt = ".csv"//Расширение по умолчани
            };

            string selectedFilePatch = string.Empty;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePatch = saveFileDialog.FileName;
            }
            if (string.IsNullOrEmpty(selectedFilePatch))
                return Result.Cancelled;

            File.WriteAllText(selectedFilePatch, wallInfo);

            return Result.Succeeded;
        }
    }
}
