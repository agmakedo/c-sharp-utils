using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Resources;
using System.Configuration;
using PISDK;
using PISDKCommon;
using PIPR;
using Utility.Log;

namespace Utility.PI
{
    class PIDataArchiveConnector
    {
        #region Data Members
        #region Private Members
        // contains the string representation of the PI Server
        private string serverName = String.Empty;
        // handle to the PI Server 
        private Server PIServer;
        // handle to the PI Server's filtered point list
        private PointList PIPointList;
        #endregion
        #endregion

        #region Constructor
        /// <summary>
        /// Default constructor that assigns the connection string to the appropriate PI Server
        /// </summary>
        public PIDataArchiveConnector(string server)
        {
            serverName = server;
        }
        #endregion 

        #region Connect To Server
        /// <summary>
        /// Attempts to establish a connection from the constructor defined PI Server
        /// </summary>
        public void Open()
        {
            try
            {
                // Create new instance of PI SDK
                PISDK.PISDK SDK = new PISDK.PISDK();
                // Connect to specified PI Server 
                Logger.Log("Connecting to: " + serverName);
                PIServer = SDK.Servers[serverName];  
                // Attempt to open connection to server with default user credentials
                Logger.Log("Logging into server as " + PIServer.DefaultUser);
                PIServer.Open();
            }
            catch (Exception ex) 
            {
                throw new System.Exception(ex.Message);
            }
        }


        /// <summary>
        /// Checks PIServer object to see if a valid connection to the specified PI Server exists
        /// </summary>
        public Boolean isConnected()
        {
            // If object is initialized, check its Connected property
            if (PIServer != null)
            {
                // returns true connection to PI Server is established
                return PIServer.Connected;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region PI Tag Collector
        /// <summary>
        /// Collect all PI tags from the specified server that match the input filter case and return a List of type String of the result
        /// </summary>
        /// <param name="filter">
        /// SQL-esque string to filter PI server and extract only desired PI tags. Filter should be modified in resource file
        /// </param>
        public List<String> CollectPIPoints(string filter) 
        {
            try
            {
                List<String> outputList = new List<string>();
                // split input filter string to perform mutliple pi point queries
                string[] filterList = filter.Split(new char[] { ',' });
                foreach (string f in filterList)
                {
                    // collect pi points based input filter
                    PIPointList = PIServer.GetPoints(f);
                    
                    Logger.Log("PI Server output " + PIPointList.Count + " tag names based on filter: " + f + ": ");
                    foreach (PISDK.PIPoint PIPoint in PIPointList)
                    {
                        // add pi point name to List object
                        outputList.Add(PIPoint.Name);
                        Logger.Log(PIPoint.Name);

                    }
                }
                return outputList;
            }
            catch (Exception ex)
            {
                throw new System.Exception(ex.Message);
            }
        }
        #endregion

        public void InsertArchiveData(string tagname, string pivalue, string timestamp)
        {
            Logger.Log("Inserting " + tagname + 
                                 " into PI Data Archive with value: " + pivalue + 
                                 " and timestamp: " + timestamp);
            PIPoint piPoint;
            piPoint = PIServer.PIPoints[tagname];
            piPoint.Data.UpdateValue(pivalue, 
                                     timestamp, 
                                     DataMergeConstants.dmReplaceDuplicates, 
                                     new PIAsynchStatus());
        }

        public void DeleteArchiveData(string tagname, string startTime, string endTime = null)
        {
            PIPoint piPoint;
            piPoint = PIServer.PIPoints[tagname];
            if (endTime == null)
            {
                Logger.Log("Deleting archive records of " + tagname + 
                                     " from PI Data Archive with timestamp: " + startTime);
                piPoint.Data.RemoveValues(startTime,
                                          startTime,
                                          DataRemovalConstants.drRemoveAll,
                                          new PIAsynchStatus());
            }
            else
            {
                Logger.Log("Deleting archive records of " + tagname + 
                                     " from PI Data Archive with start time: " + startTime + 
                                     " and end time: " + endTime);
                piPoint.Data.RemoveValues(startTime,
                                          endTime,
                                          DataRemovalConstants.drRemoveAll,
                                          new PIAsynchStatus());
            }
        }


        #region Disconnect From server
        /// <summary>
        /// Attempts to close the connection to the PI Server
        /// </summary>
        public void Close() 
        {
            try
            {
                // Attempt to close active connection to PI server
                Logger.Log("Disconnecting from " + serverName);
                PIServer.Close();
            }
            catch (Exception ex)
            {
                throw new System.Exception(ex.Message);
            }
        }
        #endregion
    }
}
