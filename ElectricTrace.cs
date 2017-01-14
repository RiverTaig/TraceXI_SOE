using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.SOESupport;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.esriSystem;
using Miner.Interop;
using Miner.Framework.Trace;

namespace TraceXI_SOE
{
    class ElectricTrace
    {

        private static IGeometricNetwork _geomNet = null;
        IWorkspace _workspace = null;
        public byte[] ElectricTraceResult(IWorkspace ws,  Dictionary<string, object> inputParams, Dictionary<int,string> layerDict)
        {
            _workspace = ws;
            Random r = new Random();
            if (inputParams["traceType"].ToString().ToUpper() == "DOWNSTREAM")
            {
                if (Convert.ToBoolean(inputParams["returnByClass"]))
                {
                    int traceResultID = Convert.ToInt16( inputParams["traceResultsID"]);
                    CreateDownstreamJSON(inputParams, traceResultID, layerDict);

                    StringBuilder sb = new StringBuilder();
                    StringWriter sw = new StringWriter(sb);
                    using (JsonWriter writer = new JsonTextWriter(sw))
                    {
                        writer.WriteStartObject();
                        {
                            writer.WritePropertyName("traceResultID");
                            writer.WriteValue(traceResultID);
                        }
                        writer.WriteEndObject();
                    }
                    byte[] array = Encoding.UTF8.GetBytes(sw.ToString());
                    return array;

                    /*
                    int traceResultID = r.Next(1, 1000000);
                    Task.Run(async () => await CreateDownstreamJSON(inputParams, traceResultID, storedResults));

                    StringBuilder sb = new StringBuilder();
                    StringWriter sw = new StringWriter(sb);
                    using (JsonWriter writer = new JsonTextWriter(sw))
                    {
                        writer.WriteStartObject();
                        {
                            writer.WritePropertyName("traceResultID");
                            writer.WriteValue(traceResultID);
                        }
                        writer.WriteEndObject();
                    }
                    byte[] array = Encoding.UTF8.GetBytes(sw.ToString());
                    return array;
                     * */
                }
                else
                {
                    return CreateDownstreamJSON(inputParams, -1, layerDict);
                }
            }
            return null;
        }



        //-9168147.3,3463323
        //http://win-o7h4l8voqt9.arcfmsolution.com/arcgis104/rest/services/ArcFMMobile/SE_ElectricDistributionMobile/MapServer/exts/ArcFMMapServer/Electric%20Trace?startPoint=-9168147.3%2C3463323&traceType=Downstream&protectiveDevices=&phasesToTrace=Any&drawComplexEdges=False&includeEdges=True&includeJunctions=True&returnAttributes=True&returnGeometries=True&tolerance=100&spatialReference=&currentStatusProgID=&f=pjson
        private byte[] CreateDownstreamJSON(Dictionary<string, object> inputParams, int traceResultsID, Dictionary<int, string> layerDict)
        {

            try
            {
                IGeometricNetwork geomNet;
                DateTime dtStart;
                INetElements netEl;
                IEnumNetEID edgeEnum;
                IEnumNetEID juncEnum;
                dtStart = DateTime.Now;
                RunDownstreamTrace(inputParams["startPoint"].ToString(), out geomNet, out netEl, out edgeEnum, out juncEnum, inputParams);
                DateTime dtEnd = DateTime.Now;
                double secs = (dtEnd - dtStart).TotalSeconds;
                //Console.WriteLine("Trace Time = " + secs);
                string fieldsToGet;
                HashSet<string> fieldsToGetHash;
                GetFieldsToGet(inputParams, out fieldsToGet, out fieldsToGetHash);
                Dictionary<int, HashSet<int>> fcToOIDS = GetResultantObjectIDS(netEl, edgeEnum, juncEnum);
                dtStart = DateTime.Now;
                StringBuilder sb = new StringBuilder();
                StringWriter sw = new StringWriter(sb);

                using (JsonWriter writer = new JsonTextWriter(sw))
                {
                    writer.WriteStartObject();
                    {
                        writer.WritePropertyName("startTime");
                        writer.WriteValue(DateTime.Now.ToString());
                        writer.WritePropertyName("results");
                        writer.WriteStartArray();
                        {
                            int featureClassCounter = 1;
                            foreach (KeyValuePair<int, HashSet<int>> kvp in fcToOIDS)
                            {
                                IFeatureClass feClass = ((IFeatureClassContainer)geomNet).get_ClassByID(kvp.Key);
                                if (! IncludeFeatureClass(feClass, fieldsToGetHash))
                                {
                                    continue;
                                }
                                StringBuilder traceResultsStringBuilder = new StringBuilder();
                                StringWriter traceResultsStringWriter = new StringWriter(traceResultsStringBuilder);
                                using (JsonWriter traceResultsWriter = new JsonTextWriter(traceResultsStringWriter))
                                {
                                    if (traceResultsID > 0)
                                    {
                                        traceResultsWriter.WriteStartObject();
                                        traceResultsWriter.WritePropertyName("startTime");
                                        traceResultsWriter.WriteValue(DateTime.Now.ToString());
                                    }
                                    else
                                    {
                                        writer.WriteStartObject();
                                    } 
                                    if (layerDict.ContainsKey(feClass.FeatureClassID))
                                    {
                                        int[] oidsToGet = kvp.Value.ToArray();
                                        WriteSimpleProperties(writer, feClass, traceResultsWriter, traceResultsID, layerDict);
                                        WriteSpatialReference(writer, traceResultsWriter, traceResultsID);
                                        WriteFieldAliases(writer, feClass, fieldsToGetHash, traceResultsWriter, traceResultsID);
                                        Dictionary<int, string> fieldNameDict = new Dictionary<int, string>();
                                        WriteFields(writer, feClass, ref fieldNameDict, fieldsToGetHash, traceResultsWriter, traceResultsID);
                                        WriteFeatures(writer, feClass, oidsToGet, fieldNameDict, fieldsToGetHash,
                                            inputParams["geometriesToRetrieve"].ToString(), traceResultsWriter, traceResultsID);
                                    }
                                    if (traceResultsID > 0)
                                    {
                                        traceResultsWriter.WritePropertyName("endTime");
                                        traceResultsWriter.WriteValue(DateTime.Now.ToString());
                                        traceResultsWriter.WriteEndObject();
                                        byte[] traceResultsArray = Encoding.UTF8.GetBytes(traceResultsStringWriter.ToString());
                                        var str = System.Text.Encoding.Default.GetString(traceResultsArray);
                                        File.WriteAllText(@"C:\inetpub\wwwroot\TraceResults\" + traceResultsID.ToString() + "_" + featureClassCounter + ".json", str);
                                    }
                                    else
                                    {

                                        writer.WriteEndObject();
                                    }
                                }
                                featureClassCounter++;
                            }
                        }
                        writer.WriteEndArray();
                        writer.WritePropertyName("endTime");
                        writer.WriteValue(DateTime.Now.ToString());
                    }
                    writer.WriteEndObject();
                    //writer.WriteEnd();
                }
                byte[] array = Encoding.UTF8.GetBytes(sw.ToString());
                return array;

            }
            catch (Exception ex)
            {
                byte[] array = Encoding.UTF8.GetBytes(ex.ToString());
                return array;
            }
            finally
            {
            }
            //return null;
        }

        private static void GetFieldsToGet(Dictionary<string, object> inputParams, out string fieldsToGet, out HashSet<string> fieldsToGetHash)
        {
            fieldsToGet = inputParams["fieldsToRetrieve"].ToString();
            List<string> fieldsToGetList = fieldsToGet.Split(',').ToList();
            fieldsToGetHash = new HashSet<string>();
            for (int i = 0; i < fieldsToGetList.Count; i++)
            {
                fieldsToGetHash.Add(fieldsToGetList[i].ToUpper());
            }
        }

        private bool IncludeFeatureClass(IFeatureClass feClass, HashSet<string> fieldsToGetHash)
        {
            if (fieldsToGetHash.Contains("*"))
            {
                return true;
            }
            string feClassName = (feClass as IDataset).Name.ToUpper();
            feClassName = feClassName.Substring(1 + feClassName.LastIndexOf("."));
            if (fieldsToGetHash.Any(x => x.Contains(feClassName)))
            {
                return true;
            }
            return false;
        }

        private static Dictionary<int, HashSet<int>> GetResultantObjectIDS(INetElements netEl, IEnumNetEID edgeEnum, IEnumNetEID juncEnum)
        {
            Dictionary<int, HashSet<int>> fcToOIDS = new Dictionary<int, HashSet<int>>();//Key is ClassID, value is all OIDS in trace
            List<IEnumNetEID> enumNetEIDS = new List<IEnumNetEID> { edgeEnum, juncEnum };
            HashSet<string> classID_OIDFound = new HashSet<string>();
            foreach (IEnumNetEID enumNetEID in enumNetEIDS)
            {
                enumNetEID.Reset();
                int eid = enumNetEID.Next();
                while (eid > 0)
                {
                    int userClassID = -1; int userID = -1; int userSubID = -1;
                    netEl.QueryIDs(eid, enumNetEID.ElementType, out userClassID, out userID, out userSubID);
                    if (fcToOIDS.ContainsKey(userClassID) == false)
                    {
                        fcToOIDS.Add(userClassID, new HashSet<int>());
                    }
                    fcToOIDS[userClassID].Add(userID);
                    eid = enumNetEID.Next();
                }
            }
            return fcToOIDS;
        }

        private  string RunDownstreamTrace(string startPoint, out IGeometricNetwork geomNet, out INetElements netEl, out IEnumNetEID edgeEnum, out IEnumNetEID juncEnum, Dictionary<string,object> inputParams)
        {
            //startPoint = "-9180488.711,3460151.459";
            //startPoint = "-9180379.822,3460191.927";
            geomNet = GetGeometricNetwork("ElecGeomNet");

            int startEID = Convert.ToInt16(inputParams["startEID"]);
            if(startEID == -1)
            {
                startEID = GetStartEID(startPoint, geomNet, Convert.ToInt16(inputParams["tolerance"]));
            }
            //propSet.GetAllProperties(out names, out values1);
            Miner.Interop.SetOfPhases phaseToTrace = GetPhasesToTraceOn(inputParams);

            netEl = (INetElements)geomNet.Network;
            INetwork network = (INetwork)geomNet.Network;

            //Set up the barriers (this uses the network analysis ext, although it could be rewritten so that it did not)
            int[] barrierJuncs = new int[0];
            int[] barrierEdges = new int[0];

            //Call the trace
            IMMTracedElements tracedJunctions;
            IMMTracedElements tracedEdges;
            IMMElectricTraceSettings mmElectricTraceSettings = new MMElectricTraceSettingsClass();
            IMMElectricTracing mmElectricTracing = new MMFeederTracerClass();

            mmElectricTracing.TraceDownstream(
                geomNet, //Geometric Network
                mmElectricTraceSettings, //Trace Settings
                null, //Implementation of IMMCurrentStatus (no code in this case)
                startEID, //EID of start features
                esriElementType.esriETEdge, //Type of feature
                phaseToTrace, //Phases to trace on 
                mmDirectionInfo.establishBySourceSearch, //How to find source
                0, //Upstream neighbor (not used here)
                esriElementType.esriETNone, //Upstream neighbor type (not used here)
                barrierJuncs, // Junction barriers
                barrierEdges,//Edge barriers
                false, //Exclude open devices
                out tracedJunctions, //Resultant junctions
                out tracedEdges); //Resultant edges

            edgeEnum = NetworkHelper.GetEnumNetEID(network, tracedEdges, esriElementType.esriETEdge);
            juncEnum = NetworkHelper.GetEnumNetEID(network, tracedJunctions, esriElementType.esriETJunction);
            return startPoint;
        }

        private static SetOfPhases GetPhasesToTraceOn(Dictionary<string, object> inputParams)
        {
            string phasesToTrace = inputParams["phasesToTrace"].ToString().ToUpper();
            Miner.Interop.SetOfPhases phaseToTrace = SetOfPhases.abc;
            switch (phasesToTrace)
            {
                case "A":
                    phaseToTrace = SetOfPhases.a;
                    break;
                case "B":
                    phaseToTrace = SetOfPhases.b;
                    break;
                case "C":
                    phaseToTrace = SetOfPhases.c;
                    break;
                case "AB":
                    phaseToTrace = SetOfPhases.ab;
                    break;
                case "AC":
                    phaseToTrace = SetOfPhases.ac;
                    break;
                case "BC":
                    phaseToTrace = SetOfPhases.bc;
                    break;
                case "ABC":
                    phaseToTrace = SetOfPhases.abc;
                    break;
                case "any":
                    phaseToTrace = SetOfPhases.abc;
                    break;
            }
            return phaseToTrace;
        }

        private static void WriteFeatures(JsonWriter writer, IFeatureClass feClass, int[] oidsToGet,
            Dictionary<int, string> fieldNameDict, HashSet<string> fieldsToGetList, string geometriesToReturn, JsonWriter traceResultsWriter, int traceResultsID)
        {

            JsonWriter writerToUse = traceResultsID > 0 ? traceResultsWriter : writer;
            IFeatureCursor feCur = null;
            try
            {
                feCur = feClass.GetFeatures(oidsToGet, false);
                writerToUse.WritePropertyName("features");
                {
                    writerToUse.WriteStartArray();
                    {
                        IFeature feInTraceResults = feCur.NextFeature();
                        string feClassName= (feClass as IDataset).Name.ToUpper();
                        feClassName = feClassName.Substring(1 + feClassName.LastIndexOf("."));
                        bool needsGeometry = false;
                        if (geometriesToReturn.Contains("*") || geometriesToReturn.ToUpper().Contains(feClassName))
                        {
                            needsGeometry = true;
                        }

                        while (feInTraceResults != null)
                        {
                            writerToUse.WriteStartObject();
                            {
                                WriteAttributes(writerToUse, feInTraceResults, fieldNameDict, fieldsToGetList);
                                WriteGeometry(writerToUse, feInTraceResults, needsGeometry, traceResultsWriter, traceResultsID);
                            }
                            writerToUse.WriteEndObject();
                            feInTraceResults = feCur.NextFeature();
                        }
                    }
                    writerToUse.WriteEndArray();
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(feCur);
            }
        }

        private static void WriteGeometry(JsonWriter writer, IFeature feInTraceResults, bool needsGeometry, JsonWriter traceResultsWriter, int traceResultsID)
        {

            JsonWriter writerToUse = traceResultsID > 0 ? traceResultsWriter : writer;

            writerToUse.WritePropertyName("geometry");
            {
                writerToUse.WriteStartObject();
                {
                    if (needsGeometry)
                    {
                        if (feInTraceResults.Shape.GeometryType == esriGeometryType.esriGeometryPoint)
                        {
                            IPoint pnt = (IPoint)feInTraceResults.Shape;
                            string x = Math.Round(pnt.X, 2).ToString();
                            string y = Math.Round(pnt.Y, 2).ToString();
                            writerToUse.WritePropertyName("x"); writerToUse.WriteRawValue(x);
                            writerToUse.WritePropertyName("y"); writerToUse.WriteRawValue(y);
                        }
                        else
                        {
                            WritePaths(writer, feInTraceResults, traceResultsWriter, traceResultsID);
                        }
                    }
                }
                writerToUse.WriteEndObject();
            }
        }

        private static void WritePaths(JsonWriter writer, IFeature feInTraceResults, JsonWriter traceResultsWriter, int traceResultsID)
        {

            JsonWriter writerToUse = traceResultsID > 0 ? traceResultsWriter : writer;

            writerToUse.WritePropertyName("paths");
            writerToUse.WriteStartArray(); //Last part of multi-part polyline
            {
                writerToUse.WriteStartArray(); //Last part of polyline (which might be part of a multi-part polyline)
                {
                    IPolyline pl = feInTraceResults.Shape as IPolyline;
                    pl.Densify(-5, 1);
                    IPointCollection pointCol = (IPointCollection)pl;
                    for (int i = 0; i < pointCol.PointCount; i++)
                    {
                        writerToUse.WriteStartArray(); //Vertex on line
                        {
                            IPoint currentPoint = pointCol.get_Point(i);
                            string x = Math.Round(currentPoint.X, 2).ToString();
                            string y = Math.Round(currentPoint.Y, 2).ToString();
                            writerToUse.WriteRaw(x + "," + y);
                        }
                        writerToUse.WriteEndArray();
                    }
                }
                writerToUse.WriteEndArray();
            }
            writerToUse.WriteEndArray();
        }

        private static void WriteAttributes(JsonWriter writer, IFeature feInTraceResults, Dictionary<int, string> fieldNameDict, HashSet<string> fieldsToGetList)
        {
            string feClassName = (feInTraceResults.Class as IDataset).Name;
            writer.WritePropertyName("attributes");
            {
                writer.WriteStartObject();
                {
                    for (int i = 0; i < feInTraceResults.Fields.FieldCount; i++)
                    {
                        IField fld = feInTraceResults.Fields.get_Field(i);
                        string potentialFieldToWrite = feClassName + "." + fld.Name;
                        if (IncludeField(feInTraceResults.Class as IFeatureClass, fld, fieldsToGetList))
                        {
                            writer.WritePropertyName(fieldNameDict[i]);
                            writer.WriteValue(GetFeatureValue(feInTraceResults, i));
                        }
                    }
                }
                writer.WriteEndObject();
            }
        }

        private static void WriteFields(JsonWriter writer, IFeatureClass feClass, ref Dictionary<int, string> fieldNameDict, HashSet<string> fieldsToGetList, JsonWriter traceResultsWriter, int traceResultsID)
        {

            JsonWriter writerToUse = traceResultsID > 0 ? traceResultsWriter : writer;

            writerToUse.WritePropertyName("fields");
            {
                writerToUse.WriteStartArray();
                {
                    for (int i = 0; i < feClass.Fields.FieldCount; i++)
                    {

                            IField fld = feClass.Fields.get_Field(i);
                            if (IncludeField(feClass, fld, fieldsToGetList))
                            {
                                writerToUse.WriteStartObject();
                                {
                                    writerToUse.WritePropertyName("name"); writerToUse.WriteValue(fld.Name);
                                    fieldNameDict[i] = fld.Name;
                                    writerToUse.WritePropertyName("type"); writerToUse.WriteValue(fld.Type.ToString());
                                    writerToUse.WritePropertyName("alias"); writerToUse.WriteValue(fld.AliasName);
                                    writerToUse.WritePropertyName("length"); writerToUse.WriteValue(fld.Length);
                                }
                                writerToUse.WriteEndObject();
                            }
                    }
                }
                writerToUse.WriteEndArray();
            }
        }
        private static bool IncludeField(IFeatureClass feClass,IField field, HashSet<string> fieldsToGetList)
        {
            if (fieldsToGetList.Contains( "*"))
            {
                return true;
            }
            if (field.Name.ToUpper() == "OBJECTID")
            {
                return true;
            }
            string feClassname = (feClass as IDataset).Name.ToUpper();
            feClassname = feClassname.Substring(1 + feClassname.LastIndexOf("."));
            if (fieldsToGetList.Contains (feClassname + ".*"))
            {
                return true;
            }
            string potentialFieldToWrite = feClassname + "." + field.Name.ToUpper();
            if (fieldsToGetList.Contains(potentialFieldToWrite.ToUpper()))
            {
                return true;
            }
            return false;
        }
        private static void WriteFieldAliases(JsonWriter writer, IFeatureClass feClass, HashSet<string> fieldsToGetList, JsonWriter traceResultsWriter, int traceResultsID)
        {

            JsonWriter writerToUse = traceResultsID > 0 ? traceResultsWriter : writer;
            writerToUse.WritePropertyName("fieldAliases");
            {
                writerToUse.WriteStartObject();
                {
                    for (int i = 0; i < feClass.Fields.FieldCount; i++)
                    {
                        IField fld = feClass.Fields.get_Field(i);
                        if (IncludeField(feClass,fld,fieldsToGetList))
                        {
                            writerToUse.WritePropertyName(fld.Name);
                            writerToUse.WriteValue(fld.AliasName);
                        }
                    }
                }
                writerToUse.WriteEndObject();
            }
        }

        private static void WriteSpatialReference(JsonWriter writer,JsonWriter traceResultsWriter,int traceResultsID)
        {

            JsonWriter writerToUse = traceResultsID > 0 ? traceResultsWriter : writer;
            writerToUse.WritePropertyName("spatialReference");
            {
                writerToUse.WriteStartObject();
                {
                    writerToUse.WritePropertyName("wkid"); writerToUse.WriteValue(102100);
                    writerToUse.WritePropertyName("latestWkid"); writerToUse.WriteValue(3857);
                }
                writerToUse.WriteEndObject();
            }
        }

        private static void WriteSimpleProperties(JsonWriter writer, IFeatureClass feClass,JsonWriter traceResultsWriter,int traceResultsID, Dictionary<int,string> layerDict)
        {
            string layerInfo = layerDict[feClass.FeatureClassID];
            string layerID = layerInfo.Split(',')[0];
            string layerName = layerInfo.Split(',')[1]; 
            JsonWriter writerToUse = traceResultsID > 0 ? traceResultsWriter : writer;
            writerToUse.WritePropertyName("name"); writerToUse.WriteValue(layerName);
            writerToUse.WritePropertyName("id"); writerToUse.WriteRawValue(layerID);
            writerToUse.WritePropertyName("displayFieldName"); writerToUse.WriteValue("OBJECTID");
            writerToUse.WritePropertyName("geometryType"); writerToUse.WriteValue(feClass.ShapeType.ToString());
            writerToUse.WritePropertyName("exceededThreshold"); writerToUse.WriteValue(false);

        }
        static string GetFeatureValue(IFeature fe, int index)
        {
            try
            {
                object o = fe.get_Value(index);
                if (o != DBNull.Value && o != null)
                {
                    return o.ToString();
                }
                else
                {
                    return "null";
                }
            }
            catch
            {
                return "null";
            }
        }
        static int GetStartEID(string pointString, IGeometricNetwork geomNet, double tolerance)
        {
            IFeatureCursor feCur = null;
            try
            {
                INetElements netElements = geomNet.Network as INetElements;
                ISpatialFilter sf = new SpatialFilterClass();
                IPoint clickPoint = new PointClass();
                IProximityOperator clickPointProx = clickPoint as IProximityOperator;
                double x = Convert.ToDouble(pointString.Split(',')[0]);
                double y = Convert.ToDouble(pointString.Split(',')[1]);
                clickPoint.PutCoords(x, y);
                IGeometry sfGeom = (clickPoint as ITopologicalOperator).Buffer(tolerance);
                sf.Geometry = sfGeom;
                sf.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects;
                IEnumFeatureClass complexEdges = geomNet.get_ClassesByType(esriFeatureType.esriFTComplexEdge);
                IEnumFeatureClass simpleEdges = geomNet.get_ClassesByType(esriFeatureType.esriFTSimpleEdge);
                List<IEnumFeatureClass> edgeSimpleAndComplexEdgeFCS = new List<IEnumFeatureClass> { complexEdges, simpleEdges };
                double closestDistanceSoFar = tolerance + 1;
                int startEID = -1;
                foreach (IEnumFeatureClass enumEdgeFCS in edgeSimpleAndComplexEdgeFCS)
                {
                    enumEdgeFCS.Reset();
                    IFeatureClass fc = enumEdgeFCS.Next();
                    while (fc != null)
                    {
                        clickPoint.SpatialReference = (fc as IGeoDataset).SpatialReference;
                        sf.GeometryField = fc.ShapeFieldName;
                        feCur = fc.Search(sf, false);
                        IFeature feWithinTolerance = feCur.NextFeature();
                        while (feWithinTolerance != null)
                        {
                            double distAway = clickPointProx.ReturnDistance(feWithinTolerance.Shape);
                            if (distAway < closestDistanceSoFar)
                            {
                                closestDistanceSoFar = distAway;
                                sf.Geometry = (clickPoint as ITopologicalOperator).Buffer(closestDistanceSoFar);
                                if (feWithinTolerance is ISimpleEdgeFeature)
                                {
                                    startEID = ((ISimpleEdgeFeature)feWithinTolerance).EID;
                                }
                                else
                                {
                                    IEnumNetEID enNetEID = netElements.GetEIDs(feWithinTolerance.Class.ObjectClassID, feWithinTolerance.OID, esriElementType.esriETEdge);
                                    enNetEID.Reset();
                                    int eid = enNetEID.Next();
                                    while (eid != -1)
                                    {
                                        double distToEdge = clickPointProx.ReturnDistance(geomNet.get_GeometryForEdgeEID(eid));
                                        if (distToEdge * .999 < closestDistanceSoFar)
                                        {
                                            int oid = feWithinTolerance.OID;
                                            startEID = eid;
                                            break;
                                        }
                                        eid = enNetEID.Next();
                                    }

                                }

                            }
                            feWithinTolerance = feCur.NextFeature();
                        }
                        fc = enumEdgeFCS.Next();
                    }
                }
                return startEID;
            }
            finally
            {
                if (feCur != null)
                {
                    Marshal.FinalReleaseComObject(feCur);
                }
            }
        }


        private  IGeometricNetwork GetGeometricNetwork(string networkName)
        {
            try
            {
                if (_geomNet != null)
                {
                    return _geomNet;
                }
                IEnumDatasetName enDSName = GetWorkspace().get_DatasetNames(esriDatasetType.esriDTFeatureDataset);
                enDSName.Reset();
                IFeatureDatasetName fdsName = enDSName.Next() as IFeatureDatasetName;
                while (fdsName != null)
                {
                    fdsName = enDSName.Next() as IFeatureDatasetName;
                    IEnumDatasetName enGNName = fdsName.GeometricNetworkNames;
                    enGNName.Reset();
                    IGeometricNetworkName gnn = enGNName.Next() as IGeometricNetworkName;
                    while (gnn != null)
                    {
                        IDatasetName dsName = gnn as IDatasetName;
                        if (dsName.Name.ToUpper().Contains(networkName.ToUpper()))
                        {
                            IGeometricNetwork gn = (gnn as IName).Open() as IGeometricNetwork;
                            _geomNet = gn;
                            return gn;

                        }
                        gnn = enGNName.Next() as IGeometricNetworkName;
                    }
                }
                return null;
            }
            catch 
            {
                return null;
            }
        }

        private  IWorkspace GetWorkspace()
        {
            try
            {
                if (_workspace != null)
                {
                    return _workspace;
                }

                return _workspace;
            }
            catch (Exception ex)
            {
                //Throw the error, because we can't recover from this error!
                throw (new ApplicationException("Unable to open workspace", ex));
            }
        }
    }
}
