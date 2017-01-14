using System;
using System.Collections;
using System.Runtime.InteropServices;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.NetworkAnalysis;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.esriSystem;
using Miner.Interop;
using Miner.Framework.Trace.Utilities;
using System.Collections.Generic;

namespace TraceXI_SOE
{
    /// <summary>
    /// Helper functions for network related processing
    /// </summary>
    public class NetworkHelper : IDisposable
    {
        

        #region Constructors / Destructors

        public NetworkHelper()
        {

        }

        #endregion




        public static List<IFeature> GetFeaturesFromEIDS(IEnumNetEID enumNeteid, IGeometricNetwork geomNet, string fieldsToGet)
        {
            List<IFeature> retFeatures = new List<IFeature>();
            IEIDHelper eidHelp = new EIDHelperClass();
            eidHelp.ReturnFeatures = true;
            eidHelp.GeometricNetwork = geomNet;
            foreach (string field in fieldsToGet.Split(','))
            {
                eidHelp.AddField(field);
            }
            IEnumEIDInfo enEIDInfo = eidHelp.CreateEnumEIDInfo(enumNeteid);
            for (int i = 0; i < enEIDInfo.Count; i++)
            {
                IEIDInfo eidIn = enEIDInfo.Next();
                IFeature fe = eidIn.Feature;
                retFeatures.Add(fe);
            }
            return retFeatures;
        }
        public class ReleaseCOMReferences
        {
            private ArrayList _arrayList = new ArrayList();
            public void RegisterForRelease(object o)
            {
                _arrayList.Add(o);
            }
            public void Release()
            {
                for (int i = _arrayList.Count - 1; i > 0; i--)
                {
                    object o = _arrayList[i];
                    if (o != null)
                    {
                        Marshal.FinalReleaseComObject(o);
                    }
                }
                _arrayList.Clear();
            }
        }
        public static IEnumNetEID GetEnumNetEID(INetwork network, IMMTracedElementDeltas mmTracedElemDeltas, esriElementType elemType)
        {
            ReleaseCOMReferences relComRef = new ReleaseCOMReferences();
            try
            {
                mmTracedElemDeltas.Reset();
                IEnumNetEIDBuilder enumEIDBuilder = new EnumNetEIDArrayClass();
                relComRef.RegisterForRelease(enumEIDBuilder);
                enumEIDBuilder.Network = network;
                enumEIDBuilder.ElementType = elemType;
                for (int i = 0; i < mmTracedElemDeltas.Count; i++)
                {
                    IMMTracedElementDelta mmTracedEl = mmTracedElemDeltas.Next();
                    enumEIDBuilder.Add(mmTracedEl.EID);
                }
                return (IEnumNetEID)enumEIDBuilder;
            }
            catch 
            {
                return null;
            }
            finally
            {
                relComRef.Release();
            }
        }

        public static IEnumNetEID GetEnumNetEID(INetwork network, IMMTracedElements mmTracedElems, esriElementType elemType)
        {
            ReleaseCOMReferences relComRef = new ReleaseCOMReferences();
            try
            {
                mmTracedElems.Reset();
                IEnumNetEIDBuilder enumEIDBuilder = new EnumNetEIDArrayClass();
                relComRef.RegisterForRelease(enumEIDBuilder);
                enumEIDBuilder.Network = network;
                enumEIDBuilder.ElementType = elemType;
                for (int i = 0; i < mmTracedElems.Count; i++)
                {
                    IMMTracedElement mmTracedEl = mmTracedElems.Next();
                    enumEIDBuilder.Add(mmTracedEl.EID);
                }
                return (IEnumNetEID)enumEIDBuilder;
            }
            catch 
            {
                //_log.Error(ex.Message, ex);
                return null;
            }
            finally
            {
                relComRef.Release();
            }
        }

        



        public static void UnSetBit(int bitToUnset, ref int weight)
        {
            int x = (int)Math.Pow(2, bitToUnset);
            weight = weight | x;
            weight = weight - x;
        }




        public static int GetOIDFromEID(int EID, INetwork pNetwork, ESRI.ArcGIS.Geodatabase.esriElementType ElementType, IGeometricNetwork geomNet)
        {
            INetElements netElements = pNetwork as INetElements;
            int userClassID = 0;
            int userID = 0;
            int userSubID = 0;
            netElements.QueryIDs(EID, ElementType, out userClassID, out userID, out userSubID);
            IFeatureClassContainer feClassCon = (IFeatureClassContainer)geomNet;
            IFeatureClass esriClass = feClassCon.get_ClassByID(userClassID);
            System.Diagnostics.Trace.WriteLine(esriClass.AliasName + " OID = " + userID.ToString());
            return userID;
        }

        public static int GetOIDFromEID(int EID, INetwork pNetwork, ESRI.ArcGIS.Geodatabase.esriElementType ElementType)
        {
            INetElements netElements = pNetwork as INetElements;
            int userClassID = 0;
            int userID = 0;
            int userSubID = 0;
            netElements.QueryIDs(EID, ElementType, out userClassID, out userID, out userSubID);
            return userID;
        }



        #region IDisposable Members

        public void Dispose()
        {
        }

        #endregion
    }
}
