using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Binding = Autodesk.Revit.DB.Binding;

namespace ReSaveFloat
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        private UIApplication uiApp;
        private Application app;

        private Document docForClose = null;
        //private UIDocument uiDoc;
        string directoryFOP = "C:\\Users\\fedor.aleksandrov\\Downloads\\AE_общие параметры.txt";

        string directory = "C:\\Users\\fedor.aleksandrov\\Downloads\\Rvt_2";

        string directorySave = "C:\\Users\\fedor.aleksandrov\\Downloads\\OneDrive_1_17.08.2022\\SaveModel2\\";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.InitialDirectory = folder;
            dialog.Multiselect = false;
            dialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            DialogResult dresult = dialog.ShowDialog();
            if (dresult != DialogResult.OK)
            {
                DialogResult result = MessageBox.Show("Не выбран файлик с адреса. Продолжить со стандарными адресами?", "Проблемка!", MessageBoxButtons.OKCancel);
                if (result != DialogResult.OK)
                {
                    return Result.Cancelled;
                }
            }
            else if (dresult == DialogResult.OK)
            {
                string filePath = dialog.FileName;

                string text = System.IO.File.ReadAllText(filePath);
                List<string> textArr = text.Split('\n').ToList();
                textArr = textArr.Select(t => t.Replace("\r", String.Empty)).ToList();

                directoryFOP = textArr[0];
                directory = textArr[1];
                directorySave = textArr[2];
            }


            var allFiles = Directory.GetFiles(directory);

            var firstFiles = allFiles.Take(10).ToList();

            List<ModelPath> modelPaths = new List<ModelPath>();
            foreach (var file in firstFiles)
            {
                modelPaths.Add(ModelPathUtils.ConvertUserVisiblePathToModelPath(file));
            }

            uiApp = commandData.Application;
            //uiDoc = uiApp.ActiveUIDocument;

            app = uiApp.Application;
            app.SharedParametersFilename = directoryFOP;
            app.FailuresProcessing += DoFailureProcessing;

            int count = 0;
            foreach (var modelPath in modelPaths)
            {

                UIDocument doc = OpenFiles(modelPath);

                //UIDocument doc = uiApp.ActiveUIDocument;
                Transaction tr = new Transaction(doc.Document, "Добавление параметра");
                tr.Start();
                AddParameter(doc);
                RecParameter(doc);
                tr.Commit();
                SaveAndClose(doc);

            }


            return Result.Succeeded;
        }

        private void SaveAndClose(UIDocument uiDoc)
        {
            Directory.CreateDirectory(directorySave);
            Document doc = uiDoc.Document;
            doc.SaveAs(directorySave + doc.Title + ".rvt");
            //doc.Close();

            docForClose = doc;
            //RevitCommandId closeDoc = RevitCommandId.LookupPostableCommandId(PostableCommand.Close);
            //uiApp.PostCommand(closeDoc);
        }

        private void AddParameter(UIDocument uiDoc)
        {
            CategorySet categorySet = new CategorySet();
            categorySet.Insert(uiDoc.Document.Settings.Categories.get_Item(BuiltInCategory.OST_Doors));
            categorySet.Insert(uiDoc.Document.Settings.Categories.get_Item(BuiltInCategory.OST_Walls));

            DefinitionFile definitionFile = app.OpenSharedParameterFile();
            Definition definition = definitionFile.Groups
                .SelectMany(group => group.Definitions)
                .FirstOrDefault(x => x.Name.Equals("AE_Комментарий"));

            Binding binding = app.Create.NewInstanceBinding(categorySet);

            BindingMap map = uiDoc.Document.ParameterBindings;
            map.Insert(definition, binding, BuiltInParameterGroup.PG_TEXT);

        }

        private void RecParameter(UIDocument uiDoc)
        {
            Document doc = uiDoc.Document;
            FamilyInstance door = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Doors)
                .Cast<FamilyInstance>()
                .FirstOrDefault(x => x.Symbol.FamilyName.Contains("Стальная"));

            door.LookupParameter("AE_Комментарий").Set("коридор");
            door.Host.LookupParameter("AE_Комментарий").Set("коридор");

            List<Wall> wallsFront = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Walls)
                .Cast<Wall>()
                .Where(x => x.Name.Contains("Бетон"))  //КР_200_Бетон
                .Where(x => x.LookupParameter("AE_Комментарий").AsString() != "коридор")
                .ToList();

            wallsFront.ForEach(x => x.LookupParameter("AE_Комментарий").Set("фасад"));

            List<Wall> wallsFront2 = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Walls)
                .Cast<Wall>()
                .Where(x => x.Name.Contains("Штукатурка"))  //НО_Х_155_Штукатурка_темно - серая + 150 утеплитель
                .ToList();

            wallsFront2.ForEach(x => x.LookupParameter("AE_Комментарий").Set("фасад"));


            List<Wall> wallsBetweenFlat = new FilteredElementCollector(doc)
               .WhereElementIsNotElementType()
               .OfCategory(BuiltInCategory.OST_Walls)
               .Cast<Wall>()
               .Where(x => x.Name.Contains("Газобетон"))  //В_200_Газобетон
               .ToList();

            wallsBetweenFlat.ForEach(x => x.LookupParameter("AE_Комментарий").Set("межквартирные"));


            List<Wall> wallsInsideFlat = new FilteredElementCollector(doc)
               .WhereElementIsNotElementType()
               .OfCategory(BuiltInCategory.OST_Walls)
               .Cast<Wall>()
               .Where(x => x.Name.Contains("ПлитыПГП"))  //В_80_ПлитыПГП_Полнотелые_Обычные
               .ToList();


            wallsInsideFlat.ForEach(x => x.LookupParameter("AE_Комментарий").Set("межкомнатные"));



        }

        private UIDocument OpenFiles(ModelPath modelPath)
        {
            OpenOptions openOptions = new OpenOptions();
            UIDocument uIDoc = uiApp.OpenAndActivateDocument(modelPath, openOptions, false, new OpenCloud());
            docForClose?.Close(false);
            return uIDoc;

        }

        public void DoFailureProcessing(object sender, FailuresProcessingEventArgs args)
        {
            FailuresAccessor fa = args.GetFailuresAccessor();

            // Inside event handler, get all warnings

            IList<FailureMessageAccessor> a = fa.GetFailureMessages();

            int count = 0;

            //foreach (FailureMessageAccessor failure in a)
            //{
            //    TaskDialog.Show("Failure", failure.GetDescriptionText());
            //    fa.ResolveFailure(failure);
            //    ++count;
            //}
            //if (0 < count && args.GetProcessingResult() == FailureProcessingResult.Continue)
            //{
            //    args.SetProcessingResult(FailureProcessingResult.ProceedWithCommit);
            //}

            foreach (FailureMessageAccessor failure in a)
            {

                var sdfds = failure.GetFailureDefinitionId();
                if (failure.GetFailureDefinitionId() == BuiltInFailures.GroupFailures.AtomViolationWhenOnePlaceInstance)
                {
                    fa.DeleteWarning(failure);
                }
            }

            args.SetProcessingResult(FailureProcessingResult.Continue);
        }

    }

    public class OpenCloud : IOpenFromCloudCallback
    {
        public OpenConflictResult OnOpenConflict(OpenConflictScenario scenario)
        {
            switch (scenario)
            {
                case OpenConflictScenario.Rollback:
                    return OpenConflictResult.DiscardLocalChangesAndOpenLatestVersion;
                case OpenConflictScenario.Relinquished:
                    return OpenConflictResult.DiscardLocalChangesAndOpenLatestVersion;
                case OpenConflictScenario.OutOfDate:
                    return OpenConflictResult.DiscardLocalChangesAndOpenLatestVersion;
                case OpenConflictScenario.VersionArchived:
                    return OpenConflictResult.DiscardLocalChangesAndOpenLatestVersion;
                default:
                    return OpenConflictResult.DiscardLocalChangesAndOpenLatestVersion;
            }

        }
    }
}
