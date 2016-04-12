using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using OSIsoft.AF;
using OSIsoft.AF.Asset;
using OSIsoft.AF.PI;
using OSIsoft.AF.Time;
using OSIsoft.AF.Data;
using Utility.Log;

namespace Utility.PI
{
    class PIAssetFrameworkConnector
    {
        #region Data Members
        #region Private Members
        // used when connecting to PI AF server
        private PISystem AFServer;
        private AFDatabase AFDB;

        // used when connecting to PI DA Server
        private PIServer DAServer;    
        #endregion

        #endregion

        #region Constructor
        /// <summary>
        /// Empty constructor
        /// </summary>
        public PIAssetFrameworkConnector() { }
        #endregion

        #region Connect to PI Server
        /// <summary>
        /// Establish connection with PI AF server/database
        /// </summary>
        public void ConnectToAssetFramework(string piAFServer, string piAFDB)
        {
            AFServer = new PISystems()[piAFServer];
            try
            {
                // Attmepts to establish connection with AF Server
                Logger.Log("Establishing connection to AF Server: " + piAFServer);
                AFServer.Connect(); 
                // If successful, Specific database is specified for AFServer object
                Logger.Log("Establishing connection to " +
                                      piAFServer + "'s " +
                                      piAFDB + " database");
                // assigned AF database object to specified AF DB name
                AFDB = AFServer.Databases[piAFDB];
            }
            catch (Exception ex)
            {
                throw new System.Exception("Failed to connect to AF Server: " + piAFServer + " ERROR: " + ex.Message);
            }

        }

        /// <summary>
        /// Establish connection with PI AF server/database specified in this application's App.config file
        /// </summary>
        public void ConnectToDataArchive(string piDAServer)
        {

            try
            {
                // Attmepts to establish connection with AF Server
                Logger.Log("Establishing connection to DA Server: " + piDAServer);
                DAServer = new PIServers()[piDAServer];
                Logger.Log("Connection Successful");
            }
            catch (Exception ex)
            {
                throw new System.Exception("Failed to connect to AF Server: " + piDAServer + " ERROR: " + ex.Message);
            }

        }
        #endregion


        #region PI Data Archive C.R.U.D Methods
        public void InsertDataArchiveRecord(string pointName, object value, string timestamp)
        {
            try
            {
                //
                PIPoint piPoint = PIPoint.FindPIPoint(DAServer, pointName);
                AFValue piPointValue = new AFValue(value, new AFTime(timestamp));
                piPoint.UpdateValue(piPointValue, AFUpdateOption.Replace);
            }
            catch (Exception ex)
            {
                throw new System.Exception("Failed to insert record into server with the following error: " + ex.Message);
            }
        }

        #endregion





        #region AF Element Tree Builder
        private void ConstructElementRoot(string rootElementName)
        {
            // attempts to add root element to AF database. Statement will catch if duplicate element is found
            try
            {
                Logger.Log("Creating root element: " + rootElementName);
                AFDB.Elements.Add(rootElementName);
                CommitChanges();
            }
            catch (Exception)
            {
                Logger.Log("root element " + rootElementName + " already exists");
            }
        }
        #endregion  

        #region AF Model Commit Change
        private void CommitChanges()
        {
            Logger.Log("Applying and checking in changes to AF database");
            AFDB.ApplyChanges();
            AFDB.CheckIn();
        }
        #endregion


        /// <summary>
        /// Collect all AF elements with the specified template. Create dictionary record for each AF element. 
        /// Key = AF Element's unique ID
        /// Value = AF Element's name
        /// </summary>
        /// <param name="templateName">
        /// String representation of AF Template name
        /// </param>
        public Dictionary<string, string> FindElementsByTemplate(string templateName)
        {
            // List object to be returned to calling method
            Dictionary<string, string> returnDict = new Dictionary<string, string>();

            Logger.Log("Searching for all AF Elements with " + templateName + " template");
            AFNamedCollection<AFElement> collectedElements = AFElement.FindElementsByTemplate(AFDB, 
                                                                                              null, 
                                                                                              AFDB.ElementTemplates[templateName], 
                                                                                              true, 
                                                                                              AFSortField.Name, 
                                                                                              AFSortOrder.Ascending, 
                                                                                              10000);
            Logger.Log("Returned " + collectedElements.Count + " records: ");
            // loop through all elements found with specified AF Template and add its name to the list
            foreach (AFElement collectedElement in collectedElements)
            {
                Logger.Log(collectedElement.Name);
                returnDict.Add(collectedElement.ID.ToString(), collectedElement.Name);
            }

            return returnDict;
        }

        /// <summary>
        /// Returns the value of the specified AF attribute as a string
        /// </summary>
        /// <param name="elementID">
        /// unique identifier (GUID) of AF attribute
        /// </param>
        /// <param name="attributeName">
        /// String representation of AF attribute name
        /// </param>
        public string GetElementAttributeValue(string elementID, string attributeName)
        {
            // Find element based on unique ID, extract specific attribute's value and convert result to string
            return GetElementAttribute(elementID, attributeName).GetValue().ToString();
        }

        //GetElementAttributes
        public string[] GetElementAttributeNames(string elementID)
        {
            List<string> afAttrNames = new List<string>();

            foreach(AFAttribute afAttr in GetElementAttributes(elementID)) 
            {
                afAttrNames.Add(afAttr.Name);
            }
            // Find element based on unique ID, extract specific attribute's value and convert result to string
            return afAttrNames.ToArray<string>();
        }


        /// <summary>
        /// Returns the value of the specified AF attribute as a string
        /// </summary>
        /// <param name="elementID">
        /// unique identifier (GUID) of AF attribute
        /// </param>
        /// <param name="attributeName">
        /// String representation of AF attribute name
        /// </param>
        /// <param name="targetValue">
        /// String representation of AF Attribute value to be searched for
        /// </param>
        /// <param name="startInterval">
        /// Optional: specify day to start time range (default is -365 aka One Years Ago)
        /// </param>
        /// <param name="endInterval">
        /// Optional: specify day to end time range (default is 0 aka Today)
        /// </param>
        /// <param name="recursionLimit">
        /// Optional: specify depth of recursion (default is 5 passes with a time range dependent on startInterval and endInterval)
        /// </param>
        public string GetElementAttributeValueTimestamp(string elementID, string attributeName, string targetValue, int startInterval = -365, int endInterval = 0, int recursionLimit = 5)
        {
            if (startInterval > 0 || endInterval > 0) 
            { 
                throw new Exception("This method analyzes historical time series data. Please use negative integers when specifying desired time range"); 
            }
            // Check to see if recursion limit has been reached
            if (recursionLimit > 0)
            {
                // Instantiate TimeRange object based on current time subtracted by startInterval's day count and endInterval's day count
                AFTimeRange myTimeRange = new AFTimeRange(new AFTime(DateTime.UtcNow.AddDays(startInterval)),
                                                          new AFTime(DateTime.UtcNow.AddDays(endInterval)));
                // Loop through all AF values found within specified timeframe
                foreach (AFValue myValue in GetElementAttribute(elementID, attributeName).GetValues(myTimeRange, 0, null).Reverse<AFValue>())
                {
                    // return timestamp if targetValue has been found
                    if (myValue.Value.ToString().Equals(targetValue))
                    {
                        return myValue.Timestamp.LocalTime.ToString();
                    }
                }
                // shift time interval based on startInterval day count and repeat looping procedure. Limit recursion attempts based on specified input parameter
                return GetElementAttributeValueTimestamp(elementID, 
                                                         attributeName, 
                                                         targetValue, 
                                                         startInterval + startInterval, 
                                                         endInterval + startInterval, 
                                                         recursionLimit - 1);
            }
            else 
            {
                // returning this means that the targetValue could not be found in the AF Element's Time Series data
                return "Indefinite";
            }
        }



        public string[] GetLastValidElementAttribute(string elementID, string attributeName, string invalidValue, int startInterval = -365, int endInterval = 0, int recursionLimit = 15)
        {
            if (startInterval > 0 || endInterval > 0)
            {
                throw new Exception("This method analyzes historical time series data. Please use negative integers when specifying desired time range");
            }
            // Check to see if recursion limit has been reached
            if (recursionLimit > 0)
            {
                // Instantiate TimeRange object based on current time subtracted by startInterval's day count and endInterval's day count
                AFTimeRange myTimeRange = new AFTimeRange(new AFTime(DateTime.UtcNow.AddDays(startInterval)),
                                                          new AFTime(DateTime.UtcNow.AddDays(endInterval)));
                // Loop through all AF values found within specified timeframe
                foreach (AFValue afValue in GetElementAttribute(elementID, attributeName).GetValues(myTimeRange, 0, null).Reverse<AFValue>())
                {
                    
                    // return timestamp if targetValue has been found
                    if (!afValue.Value.ToString().Equals(invalidValue))
                    {
                        return new string[] { afValue.Value.ToString(), afValue.Timestamp.LocalTime.ToString() };
                    }
                }
                // shift time interval based on startInterval day count and repeat looping procedure. Limit recursion attempts based on specified input parameter
                return GetLastValidElementAttribute(elementID,
                                                         attributeName,
                                                         invalidValue,
                                                         startInterval + startInterval,
                                                         endInterval + startInterval,
                                                         recursionLimit - 1);
            }
            else
            {
                // returning this means that the targetValue could not be found in the AF Element's Time Series data
                return null;
            }
        }

        /// <summary>
        /// Returns a handle to the specified AFAttribute
        /// </summary>
        /// <param name="elementID">
        /// unique identifier (GUID) of AF attribute
        /// </param>
        /// <param name="attributeName">
        /// String representation of AF attribute name
        /// </param>
        private AFAttribute GetElementAttribute(string elementID, string attributeName) 
        {
            return AFElement.FindElement(AFServer, Guid.Parse(elementID)).Attributes[attributeName];
        }


        private AFAttributes GetElementAttributes(string elementID) 
        {
            return AFElement.FindElement(AFServer, Guid.Parse(elementID)).Attributes;
        }

        public string GetElementParent(string elementID)
        {
            return AFElement.FindElement(AFServer, Guid.Parse(elementID)).Parent.Name;
        }

        public string GetAttributeDataReference(string elementID, string attributeName)
        {
            try
            {
                return AFElement.FindElement(AFServer, Guid.Parse(elementID)).Attributes[attributeName].DataReference.Name.ToString();
            }
            catch (Exception)
            {
                return "";
            }
        }

        public void RemoveAttributeValue(string elementID, string attributeName)
        {
            try
            {
                PIPoint afPiPoint = (PIPoint)AFElement.FindElement(AFServer, Guid.Parse(elementID)).Attributes[attributeName].PIPoint;
                afPiPoint.UpdateValue(GetElementAttribute(elementID, attributeName).GetValue(), AFUpdateOption.Remove, AFBufferOption.BufferIfPossible);
            }
            catch (Exception)
            { }
            
        }


        #region Disconnect From AF Server
        /// <summary>
        /// Disconnect from AF Server
        /// </summary>
        public void Close()
        {
            Logger.Log("Closing connection to PI AF server...");
            AFServer.Disconnect();
        }
        #endregion
    }
}
