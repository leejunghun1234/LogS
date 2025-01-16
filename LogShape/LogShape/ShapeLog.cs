using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.IO;
using System.IO.IsolatedStorage;
using System.Security.Cryptography.X509Certificates;
using System.Windows;
using System.Windows.Forms.VisualStyles;

namespace LogShape
{
    [Transaction(TransactionMode.Manual)]
    public class ShapeLog : IExternalApplication
    {
        //변수
        public bool checkSynchronize; // synchronize 확인 -> 그게 뭔지 잘 모르겠음
        public string folderPath;  // 로그 저장할 폴더의 경로 받기
        public string tempFolderPath = "C:\\ProgramData\\Autodesk\\Revit\\temp";
        public string userId;         // 작성자의 아이디 추출
        
        public string 안녕3ㅂㅈㄷㅂㄹㅇㅎㄹㅈㄷ;

        public JArray Slog = new JArray();

        // 파일
        public string filename;       // 파일 경로 + 이름
        public string creationGUID;   // 파일 GUID
        public string filenameShort;  // 파일 이름만

        // 로그 최종
        public string jsonFile = "";  // 로그 저장할 파일 이름
        public string jsonTime = "";  // Time 파일 로그 저장할 파일 이름
        public JObject jobject = new JObject();  // 로그를 저장할 JObject

        // timeslider 용, shape 정보 저장
        // 각각의 timestamp 별로
        public JArray stlog = new JArray();      
        
        // 이게 진짜
        public List<string> stlog2 = new List<string>();  // 그 당시의 타임 로그만 저장할 List 왜 JArray 안쓰냐면 JArray는 삭제가 안됨! 왜 안되는지 모르겠음
        public JObject STlog = new JObject();             // 최종 타임 로그 JObject

        // 중복 이름 추적, 의미: 이미 한번은 설치된 적이 있다! -> 이거는 새로 만들게 아니라 계속해서 저장하고 불러와야 하는 파일인거잖아
        public List<string> elemList = new List<string>();// 현재 어떤 Element가 변형 되었는지 확인하기 용 List, 이거도 삭제를 용이하게  하기 위해서 List로 이용
        // target Category that I'll use
        public List<string> elemCatList = new List<string> // 내가 입력하고자 하는 객체 카테고리들 집합
        {
            "Walls",
            "Floors",
            "Ceilings",
            "Windows",
            "Doors",
            "Columns",
            "Structural Columns",
            "Structural Framing",
            "Stairs",
            "Railings"
        };

        // creationGUID 통해 추적 하기 위한 Dictionary 들
        public Dictionary<string, JObject> fileAndJObject = new Dictionary<string, JObject>();         // 로그 저장할 JObject
        public Dictionary<string, string> fileAndPath = new Dictionary<string, string>();              // 로그 저장할 파일 경로

        public Dictionary<string, List<string>> timeAndList = new Dictionary<string, List<string>>();  // 타임 로그 저장할 List <> stlog2 저장
        public Dictionary<string, JObject> timeAndJObject = new Dictionary<string, JObject>();         // 타임 로그 저장할 JObject <> STlog 저장
        public Dictionary<string, string> timeAndPath = new Dictionary<string, string>();              // 타임 로그 저장할 파일 경로
        public Dictionary<string, List<string>> elemListList = new Dictionary<string, List<string>>(); // ElemList 저장할 로그

        public Dictionary<string, string> fileNameList = new Dictionary<string, string>();             // 파일 이름을 저장할 리스트 <> 왜 쓰는지 확인해봐야겠는데

        public Dictionary<string, string> volumeCheckDict = new Dictionary<string, string>();
        public Dictionary<string, string> locationCheckDict = new Dictionary<string, string>();

        public Dictionary<string, Dictionary<string, string>> volumeGUID = new Dictionary<string, Dictionary<string, string>>();
        public Dictionary<string, Dictionary<string, string>> locationGUID = new Dictionary<string, Dictionary<string, string>>();


        // Application 실행되면 폴더 경로 및 EventHandler 켜주기
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                if (string.IsNullOrEmpty(folderPath)) SetLogPath(); // 파일 경로 지정해주기
                application.ControlledApplication.DocumentChanged += new EventHandler<DocumentChangedEventArgs>(DocumentChangeTracker);
                application.ControlledApplication.FailuresProcessing += new EventHandler<FailuresProcessingEventArgs>(FailureTracker);
                application.ControlledApplication.DocumentOpened += new EventHandler<DocumentOpenedEventArgs>(DocumentOpenedTracker);
                application.ControlledApplication.DocumentCreated += new EventHandler<DocumentCreatedEventArgs>(DocumentCreatedTracker);
                application.ControlledApplication.DocumentClosing += new EventHandler<DocumentClosingEventArgs>(DocumentClosingTracker);
                application.ControlledApplication.DocumentSavedAs += new EventHandler<DocumentSavedAsEventArgs>(DocumentSavedAsTracker);
                application.ControlledApplication.DocumentSaving += new EventHandler<DocumentSavingEventArgs>(DocumentSavingTracker);
            }
            catch (Exception)
            {
                return Result.Failed;
            }
            return Result.Succeeded;
        }

        // Start 될 때 켰던거 다시 꺼주기
        public Result OnShutdown(UIControlledApplication application)
        {
            try
            {
                application.ControlledApplication.DocumentChanged -= new EventHandler<DocumentChangedEventArgs>(DocumentChangeTracker);
                application.ControlledApplication.FailuresProcessing -= new EventHandler<FailuresProcessingEventArgs>(FailureTracker);
                application.ControlledApplication.DocumentOpened -= new EventHandler<DocumentOpenedEventArgs>(DocumentOpenedTracker);
                application.ControlledApplication.DocumentCreated -= new EventHandler<DocumentCreatedEventArgs>(DocumentCreatedTracker);
                application.ControlledApplication.DocumentClosing -= new EventHandler<DocumentClosingEventArgs>(DocumentClosingTracker);
                application.ControlledApplication.DocumentSavedAs -= new EventHandler<DocumentSavedAsEventArgs>(DocumentSavedAsTracker);
                application.ControlledApplication.DocumentSaving -= new EventHandler<DocumentSavingEventArgs>(DocumentSavingTracker);
            }
            catch (Exception)
            {
                return Result.Failed;
            }
            return Result.Succeeded;
        }

        // Revit 창이 열리면 json으로 해서 실시간 반영될 수 있게 저장

        // Model에 변경 사항이 생길 경우에 켜줌
        public void DocumentChangeTracker(object sender, DocumentChangedEventArgs e)
        {
            var app = sender as Autodesk.Revit.ApplicationServices.Application;
            UIApplication uiapp = new UIApplication(app);
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            creationGUID = doc.CreationGUID.ToString();
            
            jobject = fileAndJObject[creationGUID];
            jsonFile = fileAndPath[creationGUID];

            List<string> stlog2 = timeAndList[creationGUID];    // time 별의 time 로그
            JObject STlog = timeAndJObject[creationGUID];  // 전체 time log
            
            jsonTime = timeAndPath[creationGUID];  // time log 저장할 경로

            elemList = elemListList[creationGUID]; // elemList 불러오기
            volumeCheckDict = volumeGUID[creationGUID];
            locationCheckDict = locationGUID[creationGUID];

            string timestamp = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
            string filename = Path.GetFileNameWithoutExtension(doc.PathName);

            // 변경된 객체 추적 -> 여기서는 굳이 selection 된 객체만을 대상으로 할 필요 없이 전부 해야하는거지
            ICollection<ElementId> addedElements = e.GetAddedElementIds();
            ICollection<ElementId> modifiedElements = e.GetModifiedElementIds();
            ICollection<ElementId> deletedElements = e.GetDeletedElementIds();

            bool isChanged = false;
            if (addedElements != null)
            {
                try
                {
                    foreach (ElementId eid in addedElements)
                    {
                        Element elem = doc.GetElement(eid);
                        if (!checkElemPossible(doc, elem)) continue; // elemCategory 안에 해당되는 element인지 확인하고 만약 아니라면 continue

                        string eidString = $"{eid.ToString()}_1";
                        JObject addS = exportToMeshJObject(doc, elem, eidString, timestamp, "C");
                        if (addS == null) continue; // mesh 정보를 뽑을 수 없다면 continue
                        ((JArray)jobject["ShapeLog"]).Add(addS);

                        elemList.Add(eidString);
                        stlog2.Add(eidString);
                        isChanged = true;

                        if (elem.Category.Name.ToString() == "Walls")
                        {
                            Wall wall = elem as Wall;
                            string wallVolume = elem.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsValueString();
                            volumeCheckDict[eid.ToString()] = wallVolume;
                            
                            string centerPoint = getWallCenterPoint(wall);
                            locationCheckDict[eid.ToString()] = centerPoint;
                        }
                    }
                }
                catch (Exception) { }
            }

            if (modifiedElements != null)
            {
                try
                {
                    foreach (ElementId eid in modifiedElements)
                    {
                        Element elem = doc.GetElement(eid);
                        if (!checkElemPossible(doc, elem)) continue;  // elemCategory 안에 해당되는 element인지 확인하고 만약 아니라면 continue

                        if (elem.Category.Name == "Walls")
                        {
                            Wall wall = elem as Wall;

                            string wallVolume = wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsValueString();
                            bool volumeCheck = volumeCheckDict[eid.ToString()] == wallVolume;
                            
                            string centerPoint = getWallCenterPoint(wall);
                            bool locationCheck = locationCheckDict[eid.ToString()] == centerPoint;
                            if (volumeCheck && locationCheck)
                            {
                                continue;
                            }

                            // 이건 벽의 위치 혹은 volume이 변했다는 뜻
                            locationCheckDict[eid.ToString()] = centerPoint;
                            volumeCheckDict[eid.ToString()] = wallVolume;
                        }

                        int i = 1;
                        string eidString = $"{eid.ToString()}_{i}";
                        while (elemList.Contains(eidString))
                        {
                            
                            stlog2.Remove(eidString);
                            i++;
                            eidString = $"{eid.ToString()}_{i}";
                        }
                        //JObject modiS;
                        //if ((modiS = exportToMeshJObject(doc, elem, eidString, timestamp, "M")) == null) continue;
                        string modinum = eidString.Split('_')[1];
                        JObject modiS = new JObject();
                        if (modinum == "1")
                        {
                            modiS = exportToMeshJObject(doc, elem, eidString, timestamp, "C");
                        }
                        else
                        {
                            modiS = exportToMeshJObject(doc, elem, eidString, timestamp, "M");
                        }
                        if (modiS == null) continue;

                        ((JArray)jobject["ShapeLog"]).Add(modiS);
                        elemList.Add(eidString);
                        stlog2.Add(eidString);
                        isChanged = true;
                    }
                }
                catch (Exception) { }
            }
            
            if (deletedElements != null)
            {
                try
                {
                    foreach (ElementId eid in deletedElements)
                    {
                        string eidString = $"{eid.ToString()}_1";
                        int i = 1;
                        while (elemList.Contains(eidString))
                        {
                            stlog2.Remove(eidString);
                            
                            i++;
                            eidString = $"{eid.ToString()}_{i}";
                        }
                        isChanged = true;
                    }
                }
                catch (Exception) { }
                
            }
            
            // ((JArray)jobject["ShapeLog"]).Merge(Slog, new JsonMergeSettings { MergeArrayHandling = MergeArrayHandling.Concat});
            // 최신화 해주고, 지운게 있을 수 있으니까
            if (isChanged)
            {
                makeJsonFile(jsonFile, jobject);
                JArray jarray = new JArray(stlog2);
                STlog["ShapeLog"][timestamp] = jarray.DeepClone();
                makeJsonFile(jsonTime, STlog);
            }
        }

        // 별로 필요한거 같진 않아! log에도 되어있지 않는걸 보면
        public void FailureTracker(object sender, FailuresProcessingEventArgs e)
        {
            var app = sender as Autodesk.Revit.ApplicationServices.Application;
            UIApplication uiapp = new UIApplication(app);
            UIDocument uidoc = uiapp.ActiveUIDocument;
            if (uidoc != null)
            {
                Document doc = uidoc.Document;
                string user = doc.Application.Username;
                string filename = doc.PathName;
                string filenameShort = Path.GetFileNameWithoutExtension(filename);

                FailuresAccessor failuresAccessor = e.GetFailuresAccessor();
                IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();
            }
        }

        public void DocumentOpenedTracker(object sender, DocumentOpenedEventArgs e)
        {
            Document doc = e.Document;
            setProjectInfo(doc);
            fileNameList[creationGUID] = doc.PathName.ToString();
            var startTime = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
            jsonFile = Path.Combine(folderPath, startTime + $"_{creationGUID}" + $"_{doc.Title}" + ".json");
            jsonTime = Path.Combine(folderPath, startTime + $"_{creationGUID}" + $"_{doc.Title}" + "_time.json");
            
            JObject newJObject = new JObject
            {
                ["UserId"] = userId,
                ["Filename"] = filenameShort,
                ["CreationGUID"] = creationGUID,
                ["StartTime"] = startTime,
                ["EndTime"] = "",
                ["Saved"] = "False",
                ["ShapeLog"] = new JArray()
            };
            JObject newTimeJObject = new JObject
            {
                ["UserId"] = userId,
                ["Filename"] = filenameShort,
                ["CreationGUID"] = creationGUID,
                ["StartTime"] = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss"),
                ["EndTime"] = "",
                ["Saved"] = "False",
                ["ShapeLog"] = new JObject()
            };

            fileAndPath[creationGUID] = jsonFile;
            fileAndJObject[creationGUID] = newJObject;

            timeAndPath[creationGUID] = jsonTime;
            timeAndJObject[creationGUID] = newTimeJObject;

            makeJsonFile(jsonFile, newJObject);
            makeJsonFile(jsonTime, newTimeJObject);

            string elemListListPath = tempFolderPath + $"\\{creationGUID}_elemListList.json";
            string timeAndListPath = tempFolderPath + $"\\{creationGUID}_timeAndList.json";
            string volumeGUIDPath = tempFolderPath + $"\\{creationGUID}_volumeGUID.json";
            string locationGUIDPath = tempFolderPath + $"\\{creationGUID}_locationGUID.json";

            string elemListListString = File.ReadAllText(elemListListPath);
            string timeAndListString = File.ReadAllText(timeAndListPath);
            string volumeGUIDString = File.ReadAllText(volumeGUIDPath);
            string locationGUIDString = File.ReadAllText(locationGUIDPath);

            JArray elemListListJArray = JArray.Parse(elemListListString);
            JArray timeAndListJArray = JArray.Parse(timeAndListString);
            JObject volumeGUIDJObject = JObject.Parse(volumeGUIDString);
            JObject locationGUIDJObject = JObject.Parse(locationGUIDString);

            elemListList[creationGUID] = elemListListJArray.ToObject<List<string>>();
            timeAndList[creationGUID] = timeAndListJArray.ToObject<List<string>>();

            volumeGUID[creationGUID] = volumeGUIDJObject.ToObject<Dictionary<string, string>>();
            locationGUID[creationGUID] = locationGUIDJObject.ToObject<Dictionary<string, string>>();
        }

        public void DocumentCreatedTracker(object sender, DocumentCreatedEventArgs e)
        {
            Document doc = e.Document;
            setProjectInfo(doc);
            fileNameList[creationGUID] = doc.PathName.ToString();
            var startTime = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
            jsonFile = Path.Combine(folderPath, startTime + $"_{creationGUID}" + $"_{doc.Title}" + ".json");
            jsonTime = Path.Combine(folderPath, startTime + $"_{creationGUID}" + $"_{doc.Title}" + "_time.json");
            
            JObject newJObject = new JObject
            {
                ["UserId"] = userId,
                ["Filename"] = filename,
                ["CreationGUID"] = creationGUID,
                ["StartTime"] = startTime,
                ["EndTime"] = "",
                ["Saved"] = "False",
                ["ShapeLog"] = new JArray()
            };
            JObject newTimeJObject = new JObject
            {
                ["UserId"] = userId,
                ["Filename"] = filename,
                ["CreationGUID"] = creationGUID,
                ["StartTime"] = startTime,
                ["EndTime"] = "",
                ["Saved"] = "False",
                ["ShapeLog"] = new JObject()
            };

            fileAndPath[creationGUID] = jsonFile;
            fileAndJObject[creationGUID] = newJObject;

            timeAndPath[creationGUID] = jsonTime;
            timeAndJObject[creationGUID] = newTimeJObject;

            // 이 때는 처음 create 하니까 만드는게 맞는데 open을 할 때에는 불러와야지. 저장할 때 같이 이 파일 저장하고
            elemListList[creationGUID] = new List<string>();
            timeAndList[creationGUID] = new List<string>();

            volumeGUID[creationGUID] = new Dictionary<string, string>();
            locationGUID[creationGUID] = new Dictionary<string, string>();

            makeJsonFile(jsonFile, newJObject);
            makeJsonFile(jsonTime, newTimeJObject);
        }

        public void DocumentClosingTracker(object sender, DocumentClosingEventArgs e)
        {
            Document doc = e.Document;

            setProjectInfo(doc);

            jsonFile = fileAndPath[creationGUID];
            jobject = fileAndJObject[creationGUID];

            jsonTime = timeAndPath[creationGUID ];
            STlog = timeAndJObject[creationGUID];

            var endTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            jobject["EndTime"] = endTime;
            STlog["EndTime"] = endTime;

            string saved = jobject["Saved"].ToString();
            //filenameShort = Path.GetFileNameWithoutExtension(JsonFile);
            if (saved == "" || saved == "False")
            {
                makeJsonFile(jsonFile, jobject);
                makeJsonFile(jsonTime, STlog);
            }
        }

        public void DocumentSavedAsTracker(object sender, DocumentSavedAsEventArgs e)
        {
            Document doc = e.Document;

            string filename = doc.PathName;
            string filenameShort = Path.GetFileNameWithoutExtension(filename);

            string extension = getProjectInfo(doc);

            jsonFile = extension + $"_{doc.Title}_saved.json";
            jsonTime = extension + $"_{doc.Title}_time_saved.json";

            var savedTime = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");

            jobject["Filename"] = filenameShort;
            jobject["Saved"] = "True";
            jobject["EndTime"] = savedTime;

            STlog["Filename"] = filenameShort;
            STlog["Saved"] = "True";
            STlog["EndTime"] = savedTime;

            if (fileNameList[$"{doc.CreationGUID}"] != filename)
            {
                makeJsonFile(jsonFile, jobject);
                makeJsonFile(jsonTime, STlog);

                setTempPath(tempFolderPath);
            }
        }

        public void DocumentSavingTracker(object sender, DocumentSavingEventArgs e)
        {
            Document doc = e.Document;

            string extension = getProjectInfo(doc);

            jsonFile = extension + $"_{doc.Title}_saved.json";
            jsonTime = extension + $"_{doc.Title}_time_saved.json";

            var savedTime = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");

            jobject["Filename"] = filenameShort;
            jobject["Saved"] = "True";
            jobject["EndTime"] = savedTime;

            STlog["Filename"] = filenameShort;
            STlog["Saved"] = "True";
            STlog["EndTime"] = savedTime;

            makeJsonFile(jsonFile, jobject);
            makeJsonFile(jsonTime, STlog);

            setTempPath(tempFolderPath);
        }

        public void SetLogPath()
        {
            try
            {
                FileInfo fi = new FileInfo("C:\\ProgramData\\Autodesk\\Revit\\BIG_shapeLogDirectory.txt");
                if (fi.Exists)
                {
                    string logFilePath = "C:\\ProgramData\\Autodesk\\Revit";
                    string pathFile = Path.Combine(logFilePath, "BIG_shapeLogDirectory.txt");
                    using (StreamReader readtext = new StreamReader(pathFile, true))
                    {
                        string readText = readtext.ReadLine();
                        folderPath = readText;
                    }
                }
                else
                {
                    System.Windows.Forms.FolderBrowserDialog folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
                    folderBrowser.Description = "Select a folder to save Revit Modeling shape log path";
                    folderBrowser.ShowNewFolderButton = true;
                    if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        folderPath = folderBrowser.SelectedPath;
                        string LogFilePath = "C:\\ProgramData\\Autodesk\\Revit";
                        string pathFile = Path.Combine(LogFilePath, "BIG_shapeLogDirectory.txt");
                        using (StreamWriter writetext = new StreamWriter(pathFile, true))
                        {
                            writetext.WriteLine(folderPath);
                        }
                    }
                }
            }
            catch
            {
                TaskDialog.Show("안댕", "안대~");
            }
        }

        public void setTempPath(string extension)
        {
            string elemListListPath = extension + $"\\{creationGUID}_elemListList.json";
            string timeAndListPath = extension + $"\\{creationGUID}_timeAndList.json";
            string volumeGUIDPath = extension + $"\\{creationGUID}_volumeGUID.json";
            string locationGUIDPath = extension + $"\\{creationGUID}_locationGUID.json";

            JArray elemListJObject = JArray.FromObject(elemListList[creationGUID]);
            JArray timeAndListJObject = JArray.FromObject(timeAndList[creationGUID]);
            JObject volumeListJObject = JObject.FromObject(volumeGUID[creationGUID]);
            JObject locationListJObject = JObject.FromObject(locationGUID[creationGUID]);

            makeJsonFile(elemListListPath, elemListJObject);
            makeJsonFile(timeAndListPath, timeAndListJObject);
            makeJsonFile(volumeGUIDPath, volumeListJObject);
            makeJsonFile(locationGUIDPath, locationListJObject);
        }

        public void setProjectInfo(Document doc)
        {
            userId = doc.Application.Username;
            filename = doc.PathName;
            creationGUID = doc.CreationGUID.ToString();
            filenameShort = Path.GetFileNameWithoutExtension(filename);
        }

        public string getWallCenterPoint(Wall wall)
        {
            LocationCurve wallcurve = wall.Location as LocationCurve;
            Curve curve = wallcurve.Curve;
            double endX1 = curve.GetEndPoint(0).X;
            double endY1 = curve.GetEndPoint(0).Y;
            double endZ1 = curve.GetEndPoint(0).Z;
            double endX2 = curve.GetEndPoint(0).X;
            double endY2 = curve.GetEndPoint(0).Y;
            double endZ2 = curve.GetEndPoint(0).Z;

            double centerX = (endX1 + endX2) / 2;
            double centerY = (endY1 + endY2) / 2;
            double centerZ = (endZ1 + endZ2) / 2;
            string centerPoint = $"({centerX}, {centerY}, {centerZ})";
            return centerPoint;
        }

        public string getProjectInfo(Document doc)
        {
            userId = doc.Application.Username;
            string filename = doc.PathName;
            BasicFileInfo info = BasicFileInfo.Extract(filename);
            DocumentVersion v = info.GetDocumentVersion();
            string projectId = v.VersionGUID.ToString();

            jsonFile = fileAndPath[$"{doc.CreationGUID}"];
            jobject = fileAndJObject[$"{doc.CreationGUID}"];

            STlog = timeAndJObject[$"{doc.CreationGUID}"];

            string index = folderPath + "\\" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + $"_{doc.CreationGUID}";
            string extension = jsonFile.Substring(0, index.Length);

            return extension;
        }

        public bool checkElemPossible(Document doc, Element elem)
        {
            if (elem == null || doc.GetElement(elem.GetTypeId()) == null) return false;
            double? elemVolume;
            try
            {
                 elemVolume = elem.LookupParameter("Volume").AsDouble();
            }
            catch
            {
                return false;
            }
            if (elemVolume == null || elemVolume < 0.001) return false;
            if (!elemCatList.Contains(elem.Category.Name.ToString())) return false;
            return true;
        }

        public void DocumentSynchronizingTracker(object sender, DocumentSynchronizingWithCentralEventArgs e)
        {
            checkSynchronize = true;
        }

        public void DocumentSynchronizedTracker(object sender, DocumentSynchronizedWithCentralEventArgs e)
        {
            checkSynchronize = false;
            Document doc = e.Document;
            jsonFile = getProjectInfo(doc);
            var savedTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            File.WriteAllText(jsonFile, String.Empty);
            using (var streamWriter = new StreamWriter(jsonFile, true))
            using (var writer = new JsonTextWriter(streamWriter))
            {
                // 새로 쓰고 싶은거 써주면 될 듯
                jobject["endTime"] = savedTime;
                jobject["Saved"] = "True";
                var serializer = new JsonSerializer();
                serializer.Formatting = Formatting.Indented;
                serializer.Serialize(writer, jobject);
            }
        }

        public JObject exportToMeshJObject(Document doc, Element elem, string elemid, string timestamp, string CorM)
        {
            JObject job = new JObject
            {
                ["ElementId"] = elemid,
                ["Timestamp"] = timestamp,
                ["CommandType"] = CorM,
                ["ElementCategory"] = elem.Category.Name.ToString(),
                ["Meshes"] = new JArray()
            };

            Options options = new Options()
            {
                ComputeReferences = true,
                IncludeNonVisibleObjects = true
            };

            string elemCat = elem.Category.Name.ToString();

            string material = string.Empty;
            string mName = string.Empty;
            int mTransparency = 0;
            JArray mColor = new JArray();

            GeometryElement geomElem = elem.get_Geometry(options);
            
            // element가 FamilyInstance인 경우
            if (elemCat == "Windows" || elemCat == "Doors")
            {
                foreach (GeometryObject geomObj in geomElem)
                {
                    if (geomObj is GeometryInstance geomInst)
                    {
                        XYZ basisX = geomInst.Transform.BasisX;
                        XYZ basisY = geomInst.Transform.BasisY;
                        XYZ basisZ = geomInst.Transform.BasisZ;
                        XYZ origin = geomInst.Transform.Origin;

                        double originX = geomInst.Transform.Origin.X;
                        double originY = geomInst.Transform.Origin.Y;
                        double originZ = geomInst.Transform.Origin.Z;

                        foreach (GeometryObject instObj in geomInst.SymbolGeometry)
                        {
                            if (instObj is Solid geomSolid && geomSolid.Volume > 0)
                            {
                                foreach (Face face in geomSolid.Faces)
                                {
                                    Material m = doc.GetElement(face.MaterialElementId) as Material;
                                    mName = m.Name;
                                    mTransparency = m.Transparency;
                                    mColor = new JArray(m.Color.Red, m.Color.Green, m.Color.Blue);

                                    Dictionary<string, int> vertexMap = new Dictionary<string, int>();
                                    int vertexCounter = 0;

                                    Mesh mesh = face.Triangulate();
                                    if (mesh == null) continue;

                                    JObject forFamilyInstance = new JObject
                                    {
                                        ["Material"] = mName,
                                        ["Color"] = mColor,
                                        ["Transparency"] = mTransparency,
                                        ["Vertices"] = new JArray(),
                                        ["Indices"] = new JArray()
                                    };

                                    for (int i = 0; i < mesh.NumTriangles; i++)
                                    {
                                        MeshTriangle triangle = mesh.get_Triangle(i);
                                        for (int j = 0; j < 3; j++)
                                        {
                                            XYZ vertex = triangle.get_Vertex(j);
                                            string vertexKey = $"{vertex.X},{vertex.Y},{vertex.Z}";
                                            if (!vertexMap.ContainsKey(vertexKey))
                                            {
                                                vertexMap[vertexKey] = vertexCounter++;

                                                XYZ transformedVertex = origin + (vertex.X * basisX) + (vertex.Y * basisY) + (vertex.Z * basisZ);
                                                    
                                                ((JArray)forFamilyInstance["Vertices"]).Add(new JArray(
                                                    transformedVertex.X,
                                                    transformedVertex.Y,
                                                    transformedVertex.Z
                                                    ));
                                            }
                                            ((JArray)forFamilyInstance["Indices"]).Add(vertexMap[vertexKey]);
                                        }
                                    }

                                    ((JArray)job["Meshes"]).Add(forFamilyInstance);
                                }
                            }
                        }
                    }
                }
            }
            #region 일단 넘겨잇!
            //if (elem is FamilyInstance)
            //{
            //    foreach (GeometryObject geomObj in geomElem) // GeometryElement 내부를 순회
            //    {
            //        if (geomObj is GeometryInstance geomInst) // GeometryInstance인지 확인
            //        {
            //            // SymbolGeometry 순회
            //            foreach (GeometryObject instObj in geomInst.SymbolGeometry)
            //            {
            //                if (instObj is Solid geomSolid && geomSolid.Volume > 0) // Solid 타입 확인
            //                {
            //                    foreach (Face face in geomSolid.Faces)
            //                    {
            //                        Mesh mesh = face.Triangulate();
            //                        if (mesh == null) continue;

            //                        // Mesh 삼각형 순회
            //                        for (int i = 0; i < mesh.NumTriangles; i++)
            //                        {
            //                            MeshTriangle triangle = mesh.get_Triangle(i);
            //                            for (int j = 0; j < 3; j++)
            //                            {
            //                                XYZ vertex = triangle.get_Vertex(j);
            //                                string vertexKey = $"{vertex.X},{vertex.Y},{vertex.Z}";

            //                                if (!vertexMap.ContainsKey(vertexKey))
            //                                {
            //                                    vertexMap[vertexKey] = vertexCounter++;
            //                                    vertices.Add(new JArray(vertex.X, vertex.Y, vertex.Z));
            //                                }

            //                                indices.Add(vertexMap[vertexKey]);
            //                            }
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}
            #endregion
            // element가 FamilyInstance가 아닌경우
            else
            {
                JArray vertices = new JArray();
                JArray indices = new JArray();
                Dictionary<string, int> vertexMap = new Dictionary<string, int>();
                int vertexCounter = 0;
                foreach (GeometryObject geomObj in geomElem)
                {
                    if (geomObj is Solid solid && solid.Volume > 0)
                    {
                        foreach (Face face in solid.Faces)
                        {
                            if (material == string.Empty)
                            {
                                Material m = doc.GetElement(face.MaterialElementId) as Material;
                                if (m == null) continue;
                                mName = m.Name;
                                mTransparency = m.Transparency;
                                mColor = new JArray(m.Color.Red, m.Color.Green, m.Color.Blue);
                            }

                            Mesh mesh = face.Triangulate();
                            if (mesh == null) continue;

                            for (int i = 0; i < mesh.NumTriangles; i++)
                            {
                                MeshTriangle triangle = mesh.get_Triangle(i);
                                for (int j = 0; j < 3; j++)
                                {
                                    XYZ vertex = triangle.get_Vertex(j);
                                    string vertexKey = $"{vertex.X},{vertex.Y},{vertex.Z}";
                                    if (!vertexMap.ContainsKey(vertexKey))
                                    {
                                        vertexMap[vertexKey] = vertexCounter++;
                                        vertices.Add(new JArray(vertex.X, vertex.Y, vertex.Z));
                                    }
                                    indices.Add(vertexMap[vertexKey]);
                                }
                            }
                        }
                    }
                }

                // mesh 정보를 뽑아낼 수 없을 때
                if (vertices.Count == 0) return null;

                JObject forNotFamilyInstance = new JObject
                {
                    ["Material"] = mName,
                    ["Color"] = mColor,
                    ["Transparency"] = mTransparency,
                    ["Vertices"] = vertices,
                    ["Indices"] = indices
                };

                ((JArray)job["Meshes"]).Add(forNotFamilyInstance);
            }
            

            return job;
        }
        
        public void makeJsonFile(string filePath, JObject jsonfile)
        {
            File.WriteAllText(filePath, JsonConvert.SerializeObject(jsonfile, Formatting.Indented), System.Text.Encoding.UTF8);
        }

        public void makeJsonFile(string filePath, JArray jsonfile)
        {
            File.WriteAllText(filePath, JsonConvert.SerializeObject(jsonfile, Formatting.Indented), System.Text.Encoding.UTF8);
        }
    }
}
