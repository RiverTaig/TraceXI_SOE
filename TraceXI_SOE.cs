using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Collections.Specialized;

using System.Runtime.InteropServices;

using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Server;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.SOESupport;
using System.IO;

//TODO: sign the project (project properties > signing tab > sign the assembly)
//      this is strongly suggested if the dll will be registered using regasm.exe <your>.dll /codebase


namespace TraceXI_SOE
{
    [ComVisible(true)]
    [Guid("9fb8929c-8a93-4c7b-9217-ddad53449049")]
    [ClassInterface(ClassInterfaceType.None)]
    [ServerObjectExtension("MapServer",//use "MapServer" if SOE extends a Map service and "ImageServer" if it extends an Image service.
        AllCapabilities = "",
        DefaultCapabilities = "",
        Description = "",
        DisplayName = "TraceXI_SOE",
        Properties = "",
        SupportsREST = true,
        SupportsSOAP = false)]
    public class TraceXI_SOE : IServerObjectExtension, IObjectConstruct, IRESTRequestHandler
    {
        private string soe_name;

        private IPropertySet configProps;
        private IServerObjectHelper serverObjectHelper;
        private ServerLogger logger;
        private IRESTRequestHandler reqHandler;

        public TraceXI_SOE()
        {
            soe_name = this.GetType().Name;
            logger = new ServerLogger();
            reqHandler = new SoeRestImpl(soe_name, CreateRestSchema()) as IRESTRequestHandler;
        }
        private byte[] TraceOperHandler(NameValueCollection boundVariables,
                                          JsonObject operationInput,
                                              string outputFormat,
                                              string requestProperties,
                                          out string responseProperties)
        {
            responseProperties = null;
            try
            {
                var layerDict = GetClassIDToLayerIdsAndNames(_serverObject);
                Dictionary<string, object> inputParameters = new Dictionary<string, object>();

                GetInputParamsDictionary(operationInput, inputParameters);

                IWorkspace ws = GetWorkspace();

                //JsonObject result = new JsonObject();
                ElectricTrace elecTrace = new ElectricTrace();
                byte[] results = elecTrace.ElectricTraceResult(ws, inputParameters, layerDict);
                //var str = System.Text.Encoding.Default.GetString(results);
                //File.WriteAllText(@"C:\inetpub\wwwroot\TraceResults\results.txt",str);
                return results;


            }
            catch (Exception ex)
            {
                JsonObject result = new JsonObject();
                result.AddString("error", ex.ToString());

                return Encoding.UTF8.GetBytes(result.ToJson());
            }
        }
        #region IServerObjectExtension Members
        IServerObject _serverObject = null;
        public void Init(IServerObjectHelper pSOH)
        {
            serverObjectHelper = pSOH;
            _serverObject = pSOH.ServerObject;
        }

        public void Shutdown()
        {
        }

        #endregion

        #region IObjectConstruct Members

        public void Construct(IPropertySet props)
        {
            configProps = props;
        }

        #endregion

        #region IRESTRequestHandler Members

        public string GetSchema()
        {
            return reqHandler.GetSchema();
        }

        public byte[] HandleRESTRequest(string Capabilities, string resourceName, string operationName, string operationInput, string outputFormat, string requestProperties, out string responseProperties)
        {
            return reqHandler.HandleRESTRequest(Capabilities, resourceName, operationName, operationInput, outputFormat, requestProperties, out responseProperties);
        }

        #endregion

        private RestResource CreateRestSchema()
        {
            RestResource rootRes = new RestResource(soe_name, false, TraceHandler);

            RestOperation traceOper = new RestOperation("ElectricTrace",
                                                      new string[] { "startPoint", "traceType","protectiveDevices",
                                                      "phasesToTrace","drawComplexEdges", "includeEdges", "includeJunctions",
                                                      "extraInfo", "geometriesToRetrieve","tolerance",
                                                      "spatialReference","currentStatusProgID",
                                                      "fieldsToRetrieve","useModelNames",
                                                      "runInParallel","returnByClass",
                                                      "unionOnServer", "geometryPrecision","traceResultsID"},
                                                      new string[] { "json" },
                                                      TraceOperHandler);

            RestOperation traceResultsOper = new RestOperation("TraceResults",
                                                    new string[] { "traceResultsID" }, new string[] { "json" }, TraceResultsOperHandler);

            rootRes.operations.Add(traceOper);
            rootRes.operations.Add(traceResultsOper);

            return rootRes;
        }

        private byte[] TraceHandler(NameValueCollection boundVariables, string outputFormat, string requestProperties, out string responseProperties)
        {
            responseProperties = null;

            JsonObject result = new JsonObject();
            result.AddString("Compile time", "1058");

            return Encoding.UTF8.GetBytes(result.ToJson());
            
        }

        /// <summary>
        /// Returns a dictionary keyed by class id of a feauture class that is in the map.  The value is the id of the layer + ":"  + name of layer
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static Dictionary<int, string> GetClassIDToLayerIdsAndNames(IServerObject source)
        {
            Dictionary<int, string> retDict = new Dictionary<int, string>();
            IMapServer mapServer = (IMapServer)source;
            IMapServerDataAccess da = mapServer as IMapServerDataAccess;
            for (int Index = 0; Index < mapServer.MapCount; ++Index)
            {
                for(int j = 0 ; j < mapServer.GetServerInfo(mapServer.get_MapName(Index)).MapLayerInfos.Count ; j++)
                {
                    var mapLayerInfo = mapServer.GetServerInfo(mapServer.get_MapName(Index)).MapLayerInfos.get_Element(j);
                    string mapName = mapServer.get_MapName(Index);
                    object obj = da.GetDataSource(mapName, mapLayerInfo.ID);
                    if (obj is IFeatureClass)
                    {
                        var fc = obj as IFeatureClass;
                        retDict.Add(fc.FeatureClassID, mapLayerInfo.ID.ToString() + "," + mapLayerInfo.Name);
                    }
                }
            }
            return retDict;
        }

        public static IMapLayerInfo GetLayer(IServerObject source, int layerId)
        {
            if (layerId < 0)
                return (IMapLayerInfo)null;
            if (source == null)
                return (IMapLayerInfo)null;
            IMapLayerInfo mapLayerInfo = (IMapLayerInfo)null;
            IMapServer mapServer = (IMapServer)source;
            for (int Index = 0; Index < mapServer.MapCount; ++Index)
            {
                if (mapServer.GetServerInfo(mapServer.get_MapName(Index)).MapLayerInfos.Count > layerId)
                {
                    mapLayerInfo = mapServer.GetServerInfo(mapServer.get_MapName(Index)).MapLayerInfos.get_Element(layerId);
                    break;
                }
            }
            return mapLayerInfo;
        }

        private static Dictionary<string, byte[]> _storedResults = new Dictionary<string, byte[]>();



        private IWorkspace GetWorkspace()
        {

            IMapServer map = (IMapServer)_serverObject;
            IMapServer3 mapServer = (IMapServer3)_serverObject;
            IMapServerDataAccess da = mapServer as IMapServerDataAccess;

            IFeatureClass fc = (IFeatureClass)da.GetDataSource(mapServer.DefaultMapName, 0);
            IWorkspace ws = (fc as IDataset).Workspace;

            

            return ws;
        }

        private static void GetInputParamsDictionary(JsonObject operationInput, Dictionary<string, object> inputParameters)
        {
            string startPoint; string traceType; string protectiveDevices; string phasesToTrace; string drawComplexEdges;
            bool? includeEdges; bool? includeJunctions; long? startEID; string geometriesToRetrieve; long? tolerance;
            string spatialReference; string currentStatusProgID; string fieldsToRetrieve; bool? useModelNames; bool? runInParallel;
            bool? unionOnServer; long? geometryPrecision; bool? returnByClass; long? traceResultsID;

            #region Get Variables
            bool found = operationInput.TryGetString("startPoint", out startPoint);//0    (pass as blank (null) to indicate that a startEID should be used).
            //if (!found || string.IsNullOrEmpty(startPoint))
            //    throw new ArgumentNullException("startPoint");
            found = operationInput.TryGetString("traceType", out traceType);//1
            //if (!found || string.IsNullOrEmpty(traceType))
            //    throw new ArgumentNullException("traceType");
            found = operationInput.TryGetString("protectiveDevices", out protectiveDevices);//2
            found = operationInput.TryGetString("phasesToTrace", out phasesToTrace);//3
            found = operationInput.TryGetString("drawComplexEdges", out drawComplexEdges);//4
            found = operationInput.TryGetAsBoolean("includeEdges", out includeEdges);//5
            found = operationInput.TryGetAsBoolean("includeJunctions", out includeJunctions);//6
            found = operationInput.TryGetAsLong("startEID", out startEID);//7   used only if startEID is null
            found = operationInput.TryGetString("geometriesToRetrieve", out geometriesToRetrieve);//8
            found = operationInput.TryGetAsLong("tolerance", out tolerance);//9
            found = operationInput.TryGetString("spatialReference", out spatialReference);//10
            found = operationInput.TryGetString("currentStatusProgID", out currentStatusProgID);//11
            found = operationInput.TryGetString("fieldsToRetrieve", out fieldsToRetrieve);//12
            found = operationInput.TryGetAsBoolean("useModelNames", out useModelNames);//13
            found = operationInput.TryGetAsBoolean("runInParallel", out runInParallel);//14
            found = operationInput.TryGetAsBoolean("returnByClass", out returnByClass);//15
            found = operationInput.TryGetAsBoolean("unionOnServer", out unionOnServer);//16
            found = operationInput.TryGetAsLong("geometryPrecision", out geometryPrecision);//17
            found = operationInput.TryGetAsLong("traceResultsID", out traceResultsID);//18



            inputParameters.Add("startPoint", startPoint == null ? "-9176020.284,3458035.565" : startPoint);//0
            inputParameters.Add("traceType", traceType == null ? "Downstream" : traceType);//1
            inputParameters.Add("protectiveDevices", protectiveDevices);//2
            inputParameters.Add("phasesToTrace", phasesToTrace == null ? "Any" : phasesToTrace);//3
            inputParameters.Add("drawComplexEdges", drawComplexEdges == null ? false : true);//4
            inputParameters.Add("includeEdges", includeEdges == null ? true : includeEdges);//5
            inputParameters.Add("includeJunctions", includeJunctions == null ? true : includeJunctions);//6
            inputParameters.Add("startEID", startEID == null ? -1 : startEID);//7
            inputParameters.Add("geometriesToRetrieve", geometriesToRetrieve == null ? "*" : geometriesToRetrieve);//8  (all geometries is default)
            inputParameters.Add("tolerance", tolerance == null ? 30 : tolerance);//9
            inputParameters.Add("spatialReference", spatialReference);//10
            inputParameters.Add("currentStatusProgID", currentStatusProgID);//11
            inputParameters.Add("fieldsToRetrieve", fieldsToRetrieve == null ? "Transformer.FacilityID,Transformer.RatedKVA,Transformer.PhaseDesignation,Transformer.OperatingVoltage,Transformer.StreetAddress,Fuse.CustomerCount,Fuse.StreetAdress,Fuse.FacilityID,Fuse.OperatingVoltage,Fuse.Subtype,ServicePoint.ConsumptionType,ServicePoint.StreetAddress,ServicePoint.FacilityID,ServicePoint.GPSX,ServicePoint.GPSY,Switch.FacilityID,Switch.OperatingVoltage,Switch.StreetAddress,Switch.LastUser,Switch.Subtype,PriOHElectricLineSegment.NetworkLevel,PriOHElectricLineSegment.ConductorConfiguration,PriOHElectricLineSegment.MeasuredLength,PriOHElectricLineSegment.PhaseDesignation,PriOHElectricLineSegment.NeutralSize,PriUGElectricLineSegment.NetworkLevel,PriUGElectricLineSegment.NeutralMaterial,PriUGElectricLineSegment.MeasuredLength,PriUGElectricLineSegment.PhaseDesignation,PriUGElectricLineSegment.NeutralSize,SecOHElectricLineSegment.Subtype,SecOHElectricLineSegment.PhaseDesignation,SecOHElectricLineSegment.LengthSource,SecOHElectricLineSegment.FacilityID,SecOHElectricLineSegment.MeasuredLength" : fieldsToRetrieve);//12
            inputParameters.Add("useModelNames", useModelNames == null ? true : false);//13
            inputParameters.Add("runInParallel", runInParallel == null ? false : true);//14
            inputParameters.Add("returnByClass", returnByClass == null ? false : true);//15
            inputParameters.Add("unionOnServer", unionOnServer == null ? true : false);//16
            inputParameters.Add("geometryPrecision", geometryPrecision == null ? -1 : geometryPrecision);//17
            inputParameters.Add("traceResultsID", traceResultsID == null ? -1 : traceResultsID);//18

            #endregion
        }



        private byte[] TraceResultsOperHandler(NameValueCollection boundVariables,
                                                  JsonObject operationInput,
                                                      string outputFormat,
                                                      string requestProperties,
                                                  out string responseProperties)
        {
            
            responseProperties = null;
            var dict = GetClassIDToLayerIdsAndNames(_serverObject);
            JsonObject result = new JsonObject();
            foreach (KeyValuePair<int, string> kvp in dict)
            {
                result.AddString(kvp.Key.ToString(), kvp.Value);
            }
            return Encoding.UTF8.GetBytes(result.ToJson());
            
            /*
            string traceResultsID; 
            operationInput.TryGetString("traceResultsID", out traceResultsID);
            
            if (_storedResults.ContainsKey(traceResultsID) == false)
            {
                _storedResults[traceResultsID] = Encoding.UTF8.GetBytes("Not yet set");
            }

            return (_storedResults[traceResultsID]);
             */
        }
    }
}
