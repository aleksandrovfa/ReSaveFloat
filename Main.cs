using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Binding = Autodesk.Revit.DB.Binding;

namespace ReSaveFamily
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        private UIApplication uiApp;
        private Application app;

        private UIDocument doc;
        private Document docForClose = null;
        //private UIDocument uiDoc;
        //string directoryFOP = "C:\\Users\\fedor.aleksandrov\\Downloads\\AE_общие параметры.txt";

        string directory = @"C:\Users\fedor.aleksandrov\Desktop\пересохранение семейств\Исходные\";

        string directorySave = @"C:\Users\fedor.aleksandrov\Desktop\пересохранение семейств\Пересохраненные\";
        int modelsCount = 8;
        bool modelsCountAll = true; 

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
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

                    //directoryFOP = textArr[0];
                    directory = textArr[0];
                    directorySave = textArr[1];
                    modelsCountAll = (textArr[2] == "all");
                    if (!modelsCountAll)
                    {
                        modelsCount = Convert.ToInt32(textArr[2]);
                    }
                }


                Debug.Listeners.Clear();
                Debug.Listeners.Add(new RbsLogger.Logger("ReSaveFamily"));
                Debug.WriteLine(directory);
                Debug.WriteLine(directorySave);

                var allFiles = Directory.GetFiles(directory, "*.rfa", SearchOption.AllDirectories);

                Debug.WriteLine("Всего моделей:" + allFiles.Count());
                if (modelsCountAll)
                {
                    Debug.WriteLine("Будут выгружены все модели");
                }
                else
                {
                    Debug.WriteLine($"Будет выгружены {modelsCount} модели");
                }

                List<ModelPath> modelPaths = new List<ModelPath>();
                int countModelsNow = 0;
                foreach (var file in allFiles)
                {
                    if (modelsCountAll)
                    {
                        modelPaths.Add(ModelPathUtils.ConvertUserVisiblePathToModelPath(file));
                    }
                    else
                    {
                        if(countModelsNow < modelsCount)
                        {
                            modelPaths.Add(ModelPathUtils.ConvertUserVisiblePathToModelPath(file));
                            countModelsNow++;
                        }
                    }
                }

                uiApp = commandData.Application;
                //uiDoc = uiApp.ActiveUIDocument;

                app = uiApp.Application;
                //app.SharedParametersFilename = directoryFOP;
                //app.FailuresProcessing += DoFailureProcessing;

                int count = 0;
                foreach (var modelPath in modelPaths)
                {
                    Debug.WriteLine("");
                    count++;
                    Debug.WriteLine($"ПОРЯДКОВЫЙ НОМЕР СЕМЕЙСТВА: {count.ToString()}");
                    Debug.WriteLine(ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath));
                    try
                    {
                        doc = OpenFiles(modelPath);
                        Debug.WriteLine("!!!МАТЕРИАЛЫ ДО УДАЛЕНИЯ:");
                        WriteMaterials(doc);


                        Transaction tr = new Transaction(doc.Document, "Удаление материалов");
                        tr.Start();
                        DeleteMaterials(doc);
                        tr.Commit();

                        Debug.WriteLine("!!!МАТЕРИАЛЫ ПОСЛЕ УДАЛЕНИЯ:");
                        WriteMaterials(doc);

                        SaveAndClose(doc);
                        
                    }
                    catch (Exception ex)
                    {

                        Debug.WriteLine(ex.ToString());
                    }

                   

                }
                Debug.WriteLine("Конец");
                MessageBox.Show("Пересохранение закончилось. \n Надеюсь что всё прошло удачно))))", "Конец");
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.ToString());
                MessageBox.Show(e.Message + e.StackTrace, "Что то поломалось... Скоро починим");
                return Result.Failed;
            }




            return Result.Succeeded;
        }

        private void WriteMaterials(UIDocument doc)
        {
            List<Material> materials = new FilteredElementCollector(doc.Document)
                .OfClass(typeof(Material))
                .Cast<Material>()
                .ToList();

            foreach (var material in materials)
            {
                Debug.WriteLine(material.Name);
            }
           
        }

        private void DeleteMaterials(UIDocument doc)
        {
            
            List<Material> materials = new FilteredElementCollector(doc.Document)
                .OfClass(typeof(Material))
                .Cast<Material>()
                .ToList();


            List<ElementId> listIdMaterial = new List<ElementId>();

            List<FamilyParameter> familyParameters = doc.Document.FamilyManager.GetParameters()
                .Where(x => x.StorageType == StorageType.ElementId)
                .ToList();

            foreach (var fmparem in familyParameters)
            {
                foreach (Parameter param in fmparem.AssociatedParameters)
                {
                    listIdMaterial.Add(param.AsElementId());
                }
                
            }

            var listIdMaterialDistinct = listIdMaterial.Distinct().ToList();
            
            materials = materials
                .Where(x =>!listIdMaterialDistinct.Contains(x.Id))
                .Where(x =>!x.Name.StartsWith("ETL_")).ToList();

            doc.Document.Delete(materials.Select(x => x.Id).ToList());

            


        }

        private void SaveAndClose(UIDocument uiDoc)
        {
            Directory.CreateDirectory(directorySave);
            Document doc = uiDoc.Document;
            string pathSource = doc.PathName;
            string path = pathSource.Replace(directory, directorySave);
            string pathWithoutName = path.Replace($@"\{doc.Title}.rfa", "");
            if (!Directory.Exists(pathWithoutName))
            {
                Directory.CreateDirectory(pathWithoutName);
            }

            SaveAsOptions options = new SaveAsOptions();
            options.OverwriteExistingFile = true;
            
            doc.SaveAs(path, options);
            //doc.Close();

            docForClose = doc;
            //RevitCommandId closeDoc = RevitCommandId.LookupPostableCommandId(PostableCommand.Close);
            //uiApp.PostCommand(closeDoc);
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
                    Debug.WriteLine("Ошибка группы");
                    Debug.WriteLine(failure.GetDescriptionText());
                }
                else if (failure.GetFailureDefinitionId() == BuiltInFailures.GeneralFailures.ErrorInSymbolFamilyResolved)
                {
                    fa.DeleteWarning(failure);
                    Debug.WriteLine("Ошибка семейства");
                    Debug.WriteLine(failure.GetDescriptionText());
                }
                else
                {
                    fa.DeleteWarning(failure);
                    Debug.WriteLine("ОШИБКА НИЖЕ НЕ ОБРАБОТАНА");
                    Debug.WriteLine(failure.GetDescriptionText());
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
                //case OpenConflictScenario.VersionArchived:
                //    return OpenConflictResult.DiscardLocalChangesAndOpenLatestVersion;
                default:
                    return OpenConflictResult.DiscardLocalChangesAndOpenLatestVersion;
            }

        }
    }
}
