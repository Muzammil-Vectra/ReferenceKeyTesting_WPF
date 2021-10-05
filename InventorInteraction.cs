using System;
using System.Windows;
using System.IO;
using Inventor;
using Environment = System.Environment;
using File = System.IO.File;

namespace ReferenceKeyTesting_WPF
{
    public struct DataPoints
    {
        public enum Entity
        { SurfaceBody, Face, Edge, ContextKey }

        public Entity EntityType { get; set; }
        public int Counter { get; set; }
        public string Name { get; set; }
        public string KeyContext { get; set; }
        public string RefKey { get; set; }

        public DataPoints(Entity entityType, int counter, string name, string keyContext, string refKey)
        {
            EntityType = entityType;
            Counter = counter;
            Name = name;
            KeyContext = keyContext;
            RefKey = refKey;
        }

    }
    public class InventorInteraction
    {
        public readonly Inventor.Application InventorApp;
        public Document ActiveDocument { get; }

        private int _edgeCounter = 0;
        private int _faceCounter = 0;
        private int _surfaceBodyCounter = 0;
        private ExcelInteraction _excelInteraction;
        private ReferenceKeyManagerClass _refKeyManagerClass = null;
        private delegate string GetReferenceKeyDelegate(dynamic entity, out string context);
        private GetReferenceKeyDelegate _getReferenceKeyDelegate;
        private bool _collectKeyContextAtEnd;
        public InventorInteraction()
        {
            try
            {
                if (InventorApp == null)
                {
                    InventorApp = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
                    ActiveDocument = InventorApp.ActiveDocument;
                }
            }
            catch (Exception ex)
            {
                MessageBoxResult result = MessageBox.Show("Please start the Inventor and load the document!");
                if (result == MessageBoxResult.OK)
                {
                    Environment.Exit(0);
                }
            }
        }
        public void CollectDataForActiveAssembly(bool collectKeyContextAtEndOnly = true)
        {
            if (ActiveDocument is AssemblyDocument)
            {
                _collectKeyContextAtEnd = collectKeyContextAtEndOnly;
                AssemblyDocument theAssemblyDoc = ActiveDocument as AssemblyDocument;
                ComponentOccurrences assemblyComponentOcc = theAssemblyDoc.ComponentDefinition.Occurrences;
                _refKeyManagerClass = new ReferenceKeyManagerClass(ActiveDocument);
                _refKeyManagerClass.CreateKeyContextOnce();
                _getReferenceKeyDelegate = _refKeyManagerClass.GetReferenceKeyOnly;
                if (!collectKeyContextAtEndOnly) _getReferenceKeyDelegate = _refKeyManagerClass.GetReferenceKeyAndKeyContext;
                _excelInteraction = new ExcelInteraction();

                foreach (ComponentOccurrence componentOccurrence in assemblyComponentOcc)
                {
                    CycleComponentOccurrence(componentOccurrence);
                    if (MainWindow.Token.IsCancellationRequested) break; 
                }

                Close();
            }
        }
        public void Close()
        {
            try
            {
                if (_collectKeyContextAtEnd) SaveKeyContext();
            }
            catch (Exception ex)
            {
                Extension.CreateLog(ex);
            }
            

            _excelInteraction.CloseExcel();
        }



        private void SaveKeyContext()
        {
            string path = Environment.CurrentDirectory + @"\KeyContext.txt";
            string keyContext = _refKeyManagerClass.GetKeyContext();
            File.WriteAllText(path, String.Empty);
            File.AppendAllText(path, keyContext);
        }

        public void SaveKeyContextOnClick()
        {
            try
            {
                DataPoints dataPoints = new DataPoints()
                {
                    EntityType = DataPoints.Entity.ContextKey,
                    Counter = Math.Max(_edgeCounter + 1, _faceCounter + 1),
                    Name = "Context Key ",
                    RefKey = "",
                    KeyContext = _refKeyManagerClass.GetKeyContextOnly()
                };
                _edgeCounter += 1;
                _faceCounter += 1;
                _excelInteraction.AddDataToExcel(dataPoints);
            }
            catch (Exception e)
            {
                Extension.CreateLog(e);
            }
          
        }
        private void CycleComponentOccurrence(ComponentOccurrence componentOccurrence)
        {
            if (MainWindow.Token.IsCancellationRequested)  return; 
            _surfaceBodyCounter++;
            DataPoints data = new DataPoints
            {
                EntityType = DataPoints.Entity.SurfaceBody,
                Counter = _surfaceBodyCounter + 1,
                Name = componentOccurrence.Name,
                RefKey = "",
                KeyContext = ""
            };
            _excelInteraction.AddDataToExcel(data);

            if (componentOccurrence.SurfaceBodies.Count == 0)
            {
                try
                {
                    foreach (ComponentOccurrence subComponentOccurrence in componentOccurrence.SubOccurrences)
                    {
                        CycleComponentOccurrence(subComponentOccurrence);
                    }
                }
                catch
                {
                    // ignored
                }
            }

            foreach (SurfaceBody surfBody in componentOccurrence.SurfaceBodies)
            {
                if (MainWindow.Token.IsCancellationRequested)   return; 
                _surfaceBodyCounter++;
                data.EntityType = DataPoints.Entity.SurfaceBody;
                data.Counter = _surfaceBodyCounter + 1;
                data.Name = surfBody.Name;
                data.RefKey = _getReferenceKeyDelegate(surfBody, out string sbKeyContext);
                data.KeyContext = sbKeyContext;
                _excelInteraction.AddDataToExcel(data);
                foreach (Face face in surfBody.Faces)
                {
                    if (MainWindow.Token.IsCancellationRequested) return; 
                    _faceCounter++;
                    data.EntityType = DataPoints.Entity.Face;
                    data.Counter = _faceCounter + 1;
                    data.Name = "Face_" + _faceCounter;
                    data.RefKey = _getReferenceKeyDelegate(face, out string keyContext);
                    data.KeyContext = keyContext;
                    _excelInteraction.AddDataToExcel(data);
                }
                foreach (Edge edge in surfBody.Edges)
                {
                    if (edge.CurveType != CurveTypeEnum.kUnknownCurve)
                    {
                        if (MainWindow.Token.IsCancellationRequested)  return; 
                        _edgeCounter++;
                        data.EntityType = DataPoints.Entity.Edge;
                        data.Counter = _edgeCounter + 1;
                        data.Name = "Edge_" + _edgeCounter;
                        data.RefKey = _getReferenceKeyDelegate(edge, out string keyContext);
                        data.KeyContext = keyContext;
                        _excelInteraction.AddDataToExcel(data);
                    }
                }
            }
        }

        public bool CheckReferenceKey(string referenceKey, string keyContext)
        {
            try
            {
                ReferenceKeyManagerClass refKeyManagerClass = new ReferenceKeyManagerClass(ActiveDocument);
                int contextKey = refKeyManagerClass.LoadKeyContext(keyContext);
                object obj = refKeyManagerClass.GetEntityFromReferenceKey(referenceKey, contextKey);
                if (obj != null)
                {
                    ActiveDocument.SelectSet.Clear();
                    ActiveDocument.SelectSet.Select(obj);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Extension.CreateLog(ex);
            }
            return false;
        }
    }
}
