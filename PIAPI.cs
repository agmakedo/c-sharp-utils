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
    class PIAPI
    {
        #region Data Members
        #region Private Members
        // ENUM class determines methodology type when developering for PI
        private enum PI_TYPE { AF, DA };
        private PI_TYPE piType = PI_TYPE.AF;
        
        // used when connecting to PI AF server
        private PISystem AFServer;
        private AFDatabase AFDB;       
       
        // used when connecting to PI DA Server
        private PIServer DAServer;    
        #endregion

        #region Public Members
        public delegate void PIFunction();
        // AF Objects used to PI data specified by calling function
        public AFNamedCollectionList<AFElement> AFElementCollection;

        #endregion

        #endregion

        #region Constructor
        /// <summary>
        /// Empty constructor
        /// </summary>
        public PIAPI() { }
        #endregion

        #region Connect to PI Server
        /// <summary>
        /// Establish connection to PI AF server
        /// </summary>
        /// <param name="piAFServer">
        /// String representation of PI AF Server
        /// </param>
        public void ConnectToAssetFramework(string piAFServer)
        {            
            try
            {
                // Attmepts to establish connection with AF Server
                Logger.Log("Establishing connection to AF Server: " + piAFServer);
                AFServer = new PISystems()[piAFServer];
                piType = PI_TYPE.AF;
                Logger.Log("Connection Successful");
            }
            catch (Exception ex)
            {
                throw new System.Exception("Failed to connect to AF Server: " + piAFServer + " ERROR: " + ex.Message);
            }

        }

        // <summary>
        /// Establish connection to PI AF database
        /// </summary>
        /// <param name="piAFDatabase">
        /// String representation of PI AF database
        /// </param>
        public void ConnectToAFDatabase(string piAFDatabase)
        {
            try
            {
                Logger.Log("Connecting to AF Database: " + piAFDatabase);
                AFDB = AFServer.Databases[piAFDatabase];
                Logger.Log("Connection Successful");
            }
            catch (Exception ex)
            {
                throw new System.Exception("Failed to connect to AF Database: " + piAFDatabase + " ERROR: " + ex.Message);
            }
        }

        /// <summary>
        /// Establish connection to PI Data Archive server
        /// </summary>
        /// <param name="piDAServer">
        /// String representation of PI Data Archive Server
        /// </param>
        public void ConnectToDataArchive(string piDAServer)
        {
            try
            {
                // Attmepts to establish connection with AF Server
                Logger.Log("Establishing connection to DA Server: " + piDAServer);
                DAServer = new PIServers()[piDAServer];
                piType = PI_TYPE.DA;
                Logger.Log("Connection Successful");
            }
            catch (Exception ex)
            {
                throw new System.Exception("Failed to connect to AF Server: " + piDAServer + " " + ex.Message);
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
        /// Collect all AF Elements matching specified template name. Collection stored as publicly accessible object
        /// called AFElementCollection.
        /// Value = AF Element's name
        /// </summary>
        /// <param name="templateName">
        /// String representation of AF Template name
        /// </param>
        /// <param name="piFunction">
        /// Optional: Delegate to custom callback method (Used to provide more dynamic functionality to generalized method)
        /// </param>
        /// <param name="pageSize">
        /// Optional: Maximum amount of elements capable of being queried at a time (Default set to 10000 records/query)
        /// </param>
        public void FindElementsByTemplate(string templateName, PIFunction piFunction = null, int pageSize = 10000)
        {          
            int currentElementCount = 0; // defines the current amount of elements collected
            int totalElementCount;       // defines the total amount of elements collected

            Logger.Log("Searching for all AF Elements with " + templateName + " template");            
            do
            {
                AFElementCollection = AFElement.FindElementsByTemplate(
                                        AFDB,                                   // specify AF database
                                        null,                                   // start searching from root element
                                        AFDB.ElementTemplates[templateName],    // define AFTemplate object from input 
                                        true,                                   // include derived templates 
                                        AFSortField.Name,                       // sort by name
                                        AFSortOrder.Ascending,                  // sort name in ascending order
                                        currentElementCount,                    // starting index of element collection
                                        pageSize,                               // max amount of elements returned per call
                                        out totalElementCount);                 // total objects returned that exist in db

                // break from loop if no elements are found
                if (AFElementCollection == null) { Logger.Log("No elements found...", (int)Logger.LOGLEVEL.WARN); break; }
                // call delegate function if one is provided
                if (piFunction != null) { piFunction(); }
                // increment the current element count by the total amount returned from the previous query
                currentElementCount += totalElementCount;
            } while (currentElementCount < totalElementCount);
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

    class PIMigrator
    {
        private PIServer srcServer;
        private PIServer dstServer;

        private IEnumerable<PIPoint> srcPIPointEnumerable;

        #region Constructor
        /// <summary>
        /// Empty constructor
        /// </summary>
        public PIMigrator() {}
        #endregion

        /// <summary>
        /// Establish connection to source & destination pi servers
        /// </summary>
        /// <param name="sourcePIServer">
        /// String representation of Source PI Server
        /// </param>
        /// <param name="destinationPIServer">
        /// String representation of Destination PI Server
        /// </param>
        public void ConnectToServers(string sourcePIServer, string destinationPIServer)
        {
            try
            {
                Logger.Log("Establishing connection to source Server: " + sourcePIServer);
                srcServer = new PIServers()[sourcePIServer];
                Logger.Log("Connection Successful");

                Logger.Log("Establishing connection to destination Server: " + destinationPIServer);
                dstServer = new PIServers()[destinationPIServer];
                Logger.Log("Connection Successful");

            }
            catch (Exception ex)
            {
                throw new System.Exception("Failed to connect to Server: " + ex.Message);
            }
        }

        /// <summary>
        /// Migrate PI Points to destination server using 
        /// PI Point collection from source server based on input query and specified PI Point Attribute list
        /// </summary>
        /// <param name="piPointQuery">
        /// String representation of PI Point query. Used to collect all PI Points in question from source server
        /// </param>
        /// <param name="piPointAttributes">
        /// Optional: Contains key-value pairs of PI Point attributes to be defined during PI Point creation.
        /// If no input is provided, default attributes are assumed
        /// </param>
        public void MigratePIPoints(string piPointQuery, Dictionary<string, object> piPointAttributes = null)
        {
            try
            {
                // collect all PIPoint objects based on point search
                Logger.Log("Extracting PI Points that match the following query: " + piPointQuery);                
                srcPIPointEnumerable = PIPoint.FindPIPoints(srcServer, piPointQuery);
                Logger.Log("Found " + srcPIPointEnumerable.Count() + " PI Points matching the input query");
                // create points on destination server based on parsed PIPoint names from enumerable object
                // associate each PIPoint to be created with the attribute dict from piPointAttributes input param
                Logger.Log("Creating PI Points on " + dstServer.Name);
                dstServer.CreatePIPoints(srcPIPointEnumerable.Select(x => x.Name), piPointAttributes);
                Logger.Log("PI Point migration completed successfully");
            }
            catch (Exception ex)
            {
                throw new System.Exception("Failed to create PI Points: " + ex.Message);
            }
            //PIPoint testt = piPointEnumerable.First();
            //testt.LoadAttributes();
            //object drAttrValue;
            //IEnumerable<string> piPointAttr = testt.FindAttributeNames(null);
            //foreach (string item in piPointAttr)
            //{
            //    drAttrValue = testt.GetAttribute(item);
            //    Console.WriteLine("  {0} = '{1}'", item, drAttrValue);
            //}
        }

        public void MigratePIPointData(DateTime startTime, DateTime endTime, string filterExpression = null)
        {
            // initialize AFValues object to store pi data from source server
            AFValues piPointData;
            // initialize AFTimeRange object whose datetime range is determined by the input parameters
            AFTimeRange piPointRange = new AFTimeRange(startTime.ToUniversalTime(), endTime.ToUniversalTime());
            Logger.Log("Migrating PI Point data from " + srcServer.Name + " to " + dstServer.Name +
                       "within " + startTime.ToString() + " - " + endTime.ToString() + " range");
            try
            {
                // loop through each PIPoint collected from source server
                foreach (PIPoint piPoint in srcPIPointEnumerable)
                {
                    Logger.Log("Migrating " + piPoint.Name + " data");
                    // create new instance of AFValues object to remove stale data
                    piPointData = new AFValues();                    
                    piPointData = piPoint.RecordedValues(
                                    piPointRange,            // time range to collect pi data from
                                    AFBoundaryType.Inside,   // collect data inside the specified time range
                                    filterExpression,        // filter data based on PI Performance Equation syntax                 
                                    false);                  // exclude filtered data
                    // update PI Point values on the destination server 
                    dstServer.UpdateValues(
                                    piPointData,             // pi data collected for specified pi point + time range
                                    AFUpdateOption.Replace); // insert pi data if it doesnt exist (replace otherwise)
                }
            }
            catch (Exception ex)
            {
                throw new System.Exception("Failed to migrate PI data: " + ex.Message);
            }            
        }
    }
}
