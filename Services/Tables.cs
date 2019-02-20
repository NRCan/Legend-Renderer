using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.GeoDatabaseUI;
using GSC_Legend_Renderer;
using ESRI.ArcGIS.Carto;

namespace GSC_Legend_Renderer.Services
{
    public class Tables
    {

        #region GET methods

        /// <summary>
        /// Will return a list of all table form given workspace
        /// </summary>
        /// <returns></returns>
        public static List<ITable> GetTableListFromWorkspace(IWorkspace inputWorkspace)
        {
            //Variables
            List<ITable> outputTableList = new List<ITable>();

            //Get a list of all tables from workspace
            IEnumDataset tablesEnum = inputWorkspace.get_Datasets(esriDatasetType.esriDTTable);

            //Convert enum to list
            IDataset currentDS = tablesEnum.Next();
            while (currentDS != null)
            {
                outputTableList.Add(currentDS as ITable);

                currentDS = tablesEnum.Next();
            }

            return outputTableList;

        }

        /// <summary>
        /// Will return a list of all table form given workspace
        /// </summary>
        /// <returns></returns>
        public static List<string> GetTableNameListFromWorkspace(IWorkspace inputWorkspace)
        {
            //Variables
            List<ITable> outputTableList = GetTableListFromWorkspace(inputWorkspace);
            List<string> outputTableNameList = new List<string>();

            //Convert enum to list
            foreach (ITable tables in outputTableList)
            {
                //Cast
                IDataset tableDataset = tables as IDataset;
                outputTableNameList.Add(tableDataset.Name);
            }

            return outputTableNameList;

        }

        /// <summary>
        /// Retrieve of list of value from a table field
        /// </summary>
        /// <param name="inputWorkspace">The workspace to retrieve row information</param>
        /// <param name="tableName">Reference table name to get values from</param>
        /// <param name="fieldName">Reference to field name to get values from</param>
        /// <param name="query">Input null or a string query to refine search</param>
        /// <returns></returns>
        public static List<string> GetFieldValuesFromWorkspace(IWorkspace inputWorkspace, string tableName, string fieldName, string query)
        {
            //Variables
            List<string> TableValueList = new List<string>();
            IQueryFilter queryFilter = new QueryFilter();

            //Get table object
            ITable wantedTable = OpenTableFromWorkspace(inputWorkspace, tableName);

            //Get field index
            int fieldIndex = wantedTable.FindField(fieldName);

            //Get a search cursor within table
            if (query != null)
            {
                queryFilter.WhereClause = query;
            }
            else
            {
                queryFilter = null;
            }

            ICursor cursor = wantedTable.Search(queryFilter, true);
            IRow rows = null;
            while ((rows = cursor.NextRow()) != null)
            {
                TableValueList.Add(rows.get_Value(fieldIndex).ToString());
            }

            return TableValueList;
        }

        /// <summary>
        /// Retrieve a unique value list from a table field, in option retrieve another list from another field to use in a tag object within a control object
        /// Dico{"Main":[x1,x2,x3,...], "Tag":[y1,y2,y3,...]}
        /// </summary>
        /// <param name="inputWorkspace">The input workspace to get table from</param>
        /// <param name="uTableName">Table name to get values from</param>
        /// <param name="uFieldName">Field name to get values from</param>
        /// <param name="uQuery">An SQL query to refine search, if not needed input null</param>
        /// <param name="uTag">Bool value to get another unique list from another field, with same query.</param>
        /// <param name="uTagFieldName">Field name to get another unique list from</param>
        /// <returns></returns>
        public static Dictionary<string, List<string>> GetUniqueFieldValuesFromWorkspace(IWorkspace inputWorkspace, string uTableName, string uFieldName, string uQuery, bool uTag, string uTagFieldName)
        {
            //Variables
            Dictionary<string, List<string>> uniqueValues = new Dictionary<string, List<string>>();
            IQueryFilter uQueryFilter = new QueryFilter();
            string mainList = "Main";
            string tagList = "Tag";

            //Init. dictionnary with default keys
            uniqueValues[mainList] = new List<string>();
            if (uTag)
            {
                uniqueValues[tagList] = new List<string>();
            }

            //Get table object
            ITable uWantedTable = OpenTableFromWorkspace(inputWorkspace, uTableName);

            //Get field indexex
            int uFieldIndex = uWantedTable.FindField(uFieldName);
            int uFieldTagIndex = 0;
            if (uTag)
            {
                uFieldTagIndex = uWantedTable.FindField(uTagFieldName);
            }

            //Manage query
            if (uQuery != null)
            {
                uQueryFilter.WhereClause = uQuery;
            }
            else
            {
                uQueryFilter = null;
            }

            //Iterate through table
            ICursor uCursor = uWantedTable.Search(uQueryFilter, true);
            IRow uRows = null;
            while ((uRows = uCursor.NextRow()) != null)
            {
                //If something isn't already in list
                if (uniqueValues[mainList].Contains(uRows.get_Value(uFieldIndex).ToString()) == false)
                {
                    uniqueValues[mainList].Add(uRows.get_Value(uFieldIndex).ToString());

                    if (uTag)
                    {
                        uniqueValues[tagList].Add(uRows.get_Value(uFieldTagIndex).ToString());
                    }

                }


            }

            //Release cursor
            System.Runtime.InteropServices.Marshal.ReleaseComObject(uCursor);

            return uniqueValues;
        }

        /// <summary>
        /// Will return a dictionary with values a list of all the second entered field, PROJECT ONLY
        /// </summary>
        /// <param name="inputWorkspace">The workspace to look for the table in.</param>
        /// <param name="uTableName">The table to sort</param>
        /// <param name="uFieldName">The main field to retrieve dico keys</param>
        /// <param name="uQuery">A query in case</param>
        /// <param name="uTagFieldName">The second field name that will go in the dico value as a list</param>
        /// <returns></returns>
        public static SortedDictionary<string, List<string>> GetAllDoubleUniqueFieldValuesFromWorkspace(IWorkspace inputWorkspace, string uTableName, string uFieldName, string uTagFieldName, string uQuery)
        {
            //Variables
            SortedDictionary<string, List<string>> uniqueValues = new SortedDictionary<string, List<string>>();
            IQueryFilter uQueryFilter = new QueryFilter();

            //Get table object
            ITable uWantedTable = OpenTableFromWorkspace(inputWorkspace, uTableName);

            //Get field indexex
            int uFieldIndex = uWantedTable.FindField(uFieldName);
            int uFieldTagIndex = uWantedTable.FindField(uTagFieldName);

            //Manage query
            if (uQuery != null)
            {
                uQueryFilter.WhereClause = uQuery;
            }
            else
            {
                uQueryFilter = null;
            }

            //Iterate through table
            ICursor uCursor = uWantedTable.Search(uQueryFilter, true);
            IRow uRows = null;
            while ((uRows = uCursor.NextRow()) != null)
            {
                string key = uRows.get_Value(uFieldIndex).ToString();
                string value = uRows.get_Value(uFieldTagIndex).ToString();

                if (!uniqueValues.ContainsKey(key))
                {
                    uniqueValues[key] = new List<string>();
                    uniqueValues[key].Add(value);
                }
                else
                {
                    if (!uniqueValues[key].Contains(value))
                    {
                        uniqueValues[key].Add(value);

                        //Sort
                        uniqueValues[key].Sort();
                    }

                }

            }

            //Release cursor
            System.Runtime.InteropServices.Marshal.ReleaseComObject(uCursor);

            return uniqueValues;
        }

        /// <summary>
        /// Retrieve a value list of two field unicity from a table
        /// [(Field1, Field2), (Field1, Field2)]
        /// At the end all the list should 
        /// </summary>
        /// <param name="inputWorkspace">The input workspace to get table from</param>
        /// <param name="uTableName">Table name to get values from</param>
        /// <param name="uFieldName">The two Field name to get values from, given in a double tuple</param>
        /// <param name="uQuery">An SQL query to refine search, if not needed input null</param>
        /// <returns></returns>
        public static List<Tuple<string, string>> GetUniqueDoubleFieldValuesFromWorkspace(IWorkspace inputWorkspace, string uTableName, Tuple<string, string> uFieldNameTuple, string uQuery)
        {
            //Variables
            List<Tuple<string, string>> uniqueValues = new List<Tuple<string, string>>();
            IQueryFilter uQueryFilter = new QueryFilter();

            try
            {
                //Get table object
                ITable uWantedTable = OpenTableFromWorkspace(inputWorkspace, uTableName);

                //Get field indexes
                int uField1Index = uWantedTable.FindField(uFieldNameTuple.Item1);
                int uField2Index = uWantedTable.FindField(uFieldNameTuple.Item2);

                //Manage query
                if (uQuery != null)
                {
                    uQueryFilter.WhereClause = uQuery;
                }
                else
                {
                    uQueryFilter = null;
                }

                //Iterate through table
                ICursor uCursor = uWantedTable.Search(uQueryFilter, true);
                IRow uRows = null;
                while ((uRows = uCursor.NextRow()) != null)
                {
                    //Create tuple
                    Tuple<string, string> currentTuple = new Tuple<string, string>(uRows.get_Value(uField1Index).ToString(), uRows.get_Value(uField2Index).ToString());

                    //If something isn't already in list
                    if (!uniqueValues.Contains(currentTuple))
                    {
                        uniqueValues.Add(currentTuple);
                    }

                }

                //Release cursor
                System.Runtime.InteropServices.Marshal.ReleaseComObject(uCursor);
            }

            catch (Exception getUniqueDoubleFieldValuesFromWorkspace)
            {
                MessageBox.Show("GetUniqueDoubleFieldValuesFromWorkspace (" + Services.Exceptions.LineNumber(getUniqueDoubleFieldValuesFromWorkspace).ToString() + "):" + getUniqueDoubleFieldValuesFromWorkspace.Message);
                MessageBox.Show(uFieldNameTuple.Item1 + "[item1] + " + uFieldNameTuple.Item2 + "[item2]" + uTableName + " [tableName]");
            }
            return uniqueValues;
        }

        /// <summary>
        /// Retrieve a value list of three field unicity from a table
        /// [(Field1, Field2, Field3), (Field1, Field2, Field3)]
        /// At the end all the list should 
        /// </summary>
        /// <param name="inputWorkspace">The input workspace to get table from</param>
        /// <param name="uTableName">Table name to get values from</param>
        /// <param name="uFieldName">The three Field name to get values from, given in a triple tuple</param>
        /// <param name="uQuery">An SQL query to refine search, if not needed input null</param>
        /// <returns></returns>
        public static List<Tuple<string, string, string>> GetUniqueTripleFieldValuesFromWorkspace(IWorkspace inputWorkspace, string uTableName, Tuple<string, string, string> uFieldNameTuple, string uQuery)
        {
            //Variables
            List<Tuple<string, string, string>> uniqueValues = new List<Tuple<string, string, string>>();
            IQueryFilter uQueryFilter = new QueryFilter();

            try
            {
                //Get table object
                ITable uWantedTable = OpenTableFromWorkspace(inputWorkspace, uTableName);

                //Get field indexes
                int uField1Index = uWantedTable.FindField(uFieldNameTuple.Item1);
                int uField2Index = uWantedTable.FindField(uFieldNameTuple.Item2);
                int uField3Index = uWantedTable.FindField(uFieldNameTuple.Item3);

                //Manage query
                if (uQuery != null)
                {
                    uQueryFilter.WhereClause = uQuery;
                }
                else
                {
                    uQueryFilter = null;
                }

                //Iterate through table
                ICursor uCursor = uWantedTable.Search(uQueryFilter, true);
                IRow uRows = null;
                while ((uRows = uCursor.NextRow()) != null)
                {
                    //Create tuple
                    Tuple<string, string, string> currentTuple = new Tuple<string, string, string>(uRows.get_Value(uField1Index).ToString(), uRows.get_Value(uField2Index).ToString(), uRows.get_Value(uField3Index).ToString());

                    //If something isn't already in list
                    if (!uniqueValues.Contains(currentTuple))
                    {
                        uniqueValues.Add(currentTuple);
                    }

                }

                //Release cursor
                System.Runtime.InteropServices.Marshal.ReleaseComObject(uCursor);
            }

            catch (Exception getUniqueTripleFieldValuesFromWorkspace)
            {
                MessageBox.Show("GetUniqueTripleFieldValuesFromWorkspace (" + Services.Exceptions.LineNumber(getUniqueTripleFieldValuesFromWorkspace).ToString() + "):" + getUniqueTripleFieldValuesFromWorkspace.Message);
                MessageBox.Show(uFieldNameTuple.Item1 + "[item1] + " + uFieldNameTuple.Item2 + "[item2]" + uFieldNameTuple.Item3 + "[item3]" + uTableName + " [tableName]");
            }

            return uniqueValues;
        }

        /// <summary>
        /// Retrieve a value list of three field unicity from a table, first tuple item will be used as key
        /// {item1:[item2, item3]}
        /// </summary>
        /// <param name="inputWorkspace">The workspace to get table from</param>
        /// <param name="uTableName">Table name to get values from</param>
        /// <param name="uFieldNameTuple">The three Field name to get values from, given in a triple tuple</param>
        /// <param name="uQuery">An SQL query to refine search, if not needed input null</param>
        /// <returns></returns>
        public static Dictionary<string, Tuple<string, string>> GetUniqueDicoTripleFieldValuesFromWorkspace(IWorkspace inputWorkspace, string uTableName, Tuple<string, string, string> uFieldNameTuple, string uQuery)
        {
            //Variables
            Dictionary<string, Tuple<string, string>> uniqueValues = new Dictionary<string, Tuple<string, string>>();
            IQueryFilter uQueryFilter = new QueryFilter();

            try
            {
                //Get table object
                ITable uWantedTable = OpenTableFromWorkspace(inputWorkspace, uTableName);

                //Get field indexes
                int uField1Index = uWantedTable.FindField(uFieldNameTuple.Item1);
                int uField2Index = uWantedTable.FindField(uFieldNameTuple.Item2);
                int uField3Index = uWantedTable.FindField(uFieldNameTuple.Item3);

                //Manage query
                if (uQuery != null)
                {
                    uQueryFilter.WhereClause = uQuery;
                }
                else
                {
                    uQueryFilter = null;
                }

                //Iterate through table
                ICursor uCursor = uWantedTable.Search(uQueryFilter, true);
                IRow uRows = null;
                while ((uRows = uCursor.NextRow()) != null)
                {
                    //Current key
                    string currentKey = uRows.get_Value(uField1Index).ToString();

                    //Create tuple
                    Tuple<string, string> currentTuple = new Tuple<string, string>(uRows.get_Value(uField2Index).ToString(), uRows.get_Value(uField3Index).ToString());

                    //If something isn't already in list
                    if (!uniqueValues.ContainsKey(currentKey))
                    {
                        uniqueValues[currentKey] = currentTuple;
                    }

                }

                //Release cursor
                System.Runtime.InteropServices.Marshal.ReleaseComObject(uCursor);

            }

            catch (Exception getUniqueDicoTripleFieldValuesFromWorkspaceEx)
            {
                MessageBox.Show(getUniqueDicoTripleFieldValuesFromWorkspaceEx.StackTrace);
            }
            return uniqueValues;
        }

        /// <summary>
        /// Will return a dictionnary with keys as first field unique values and and keys values as second field first encountered value
        /// </summary>
        /// <param name="uTableName">The table to extract data from</param>
        /// <param name="uFieldName">First field name that values will become dico keys</param>
        /// <param name="u2FieldName">Second field name that values will become dico values</param>
        /// <param name="uQuery">A query to filter table, this can be null</param>
        /// <returns></returns>
        public static Dictionary<string, string> GetUniqueDicoValuesFromWorkspace(IWorkspace inputWorkspace, string uTableName, string uFieldName, string u2FieldName, string uQuery)
        {
            //Variables
            Dictionary<string, string> uniqueValues = new Dictionary<string, string>();
            IQueryFilter uQueryFilter = new QueryFilter();

            //Get table object
            ITable uWantedTable = OpenTableFromWorkspace(inputWorkspace, uTableName);

            //Get field indexex
            int uFieldIndex = uWantedTable.FindField(uFieldName);
            int u2FieldIndex = uWantedTable.FindField(u2FieldName);

            //Manage query
            if (uQuery != null)
            {
                uQueryFilter.WhereClause = uQuery;
            }
            else
            {
                uQueryFilter = null;
            }

            //Iterate through table
            ICursor uCursor = uWantedTable.Search(uQueryFilter, true);
            IRow uRows = null;
            while ((uRows = uCursor.NextRow()) != null)
            {
                if (!uniqueValues.ContainsKey(uRows.get_Value(uFieldIndex).ToString()))
                {
                    uniqueValues[uRows.get_Value(uFieldIndex).ToString()] = uRows.get_Value(u2FieldIndex).ToString();
                }

            }

            //Release cursor
            System.Runtime.InteropServices.Marshal.ReleaseComObject(uCursor);

            return uniqueValues;

        }

        /// <summary>
        /// Will return a cursor object to iterate through a table
        /// </summary>
        /// <param name="cursorType">Cursor type: Update, Insert or Search</param>
        /// <param name="query">A query to filter search or update cursors</param>
        /// <param name="featureName">the desire table name to get the cursor from</param>
        /// <returns></returns>
        public static ICursor GetTableCursorFromWorkspace(IWorkspace inputWorkspace, string cursorType, string query, string tableName)
        {
            //Main variable
            ICursor getTableCursor = null;

            //Get required feature class object
            ITable getTable = OpenTableFromWorkspace(inputWorkspace, tableName);

            //Build a query filter for update cursor
            IQueryFilter queryFilter = new QueryFilter();

            //Update filter with where clause
            if (query != null)
            {
                queryFilter.WhereClause = query;
            }
            else
            {
                queryFilter = null;
            }

            //Access inside feature class with a cursor (an update one in case there is a need to recalculate values)
            if (cursorType == "Update")
            {
                try
                {
                    getTableCursor = getTable.Update(queryFilter, false);
                }
                catch (Exception e)
                {

                    MessageBox.Show(e.StackTrace);
                }

            }

            else if (cursorType == "Search")
            {
                getTableCursor = getTable.Search(queryFilter, true);
            }

            else if (cursorType == "Insert")
            {
                getTableCursor = getTable.Insert(true);
            }
            else
            {
                getTableCursor = getTable.Search(queryFilter, true);
            }

            return getTableCursor;
        }

        /// <summary>
        /// Will calculate row count from given table name, within a given workspace
        /// </summary>
        /// <param name="inputWorkspace">The workspace to get the table count from</param>
        /// <param name="queryFilter">An sql query to filter the count</param>
        /// <param name="tableToCount">The table to count row from</param>
        /// <returns></returns>
        public static int GetRowCountFromWorkspace(IWorkspace inputWorkspace, string tableToCount, string queryFilter)
        {
            //Get row count for input table
            ITable getTable = OpenTableFromWorkspace(inputWorkspace, tableToCount);

            //Build a query filter if needed
            IQueryFilter countFilter = new QueryFilter();
            if (queryFilter != null)
            {
                countFilter.WhereClause = queryFilter;
            }


            return getTable.RowCount(countFilter);
        }

        /// <summary>
        /// Will return a string list containing all the fields from given table
        /// </summary>
        /// <param name="inputFC">The table that contains the fields</param>
        /// <param name="alias">Speficify true if field name has to be alias instead of real name</param>
        /// <returns></returns>
        public static List<string> GetFieldList(ITable inputTable, bool alias, IStandaloneTable inSTL = null)
        {
            //Variables
            List<string> fieldList = new List<string>();
            int fieldCount = 0;

            //Get the fields object
            IFields inFields = inputTable.Fields;

            //Iterate through collection
            while (fieldCount < inFields.FieldCount)
            {
                if (alias)
                {
                    fieldList.Add(inFields.Field[fieldCount].AliasName);
                }
                else
                {
                    fieldList.Add(inFields.Field[fieldCount].Name);
                }

                fieldCount++;
            }

            //Add fields from any joined tables
            if (inSTL !=null)
            {
                ITableFields tFields = inSTL as ITableFields;
                for (int f = 0; f < tFields.FieldCount; f++)
                {
                    string tFieldName = string.Empty;
                    if (alias)
                    {
                        tFieldName = tFields.Field[f].AliasName;
                    }
                    else
                    {
                        tFieldName = tFields.Field[f].Name;
                    }

                    if (!fieldList.Contains(tFieldName))
                    {
                        fieldList.Add(tFieldName);
                    }
                }
            }

            return fieldList;
        }

        #endregion

        #region OPEN methods

        /// <summary>
        /// To open a table from a workspace
        /// </summary>
        /// <param name="inputWorkspace">Reference to a workspace object</param>
        /// <param name="inputTable">Reference table name to open</param>
        /// <returns></returns>
        public static ITable OpenTableFromWorkspace(IWorkspace inputWorkspace, string inputTable)
        {
            //Access workspace
            IFeatureWorkspace getWorkspace = (IFeatureWorkspace)inputWorkspace;

            ITable outTable = getWorkspace.OpenTable(inputTable);

            Services.ObjectManagement.ReleaseObject(getWorkspace);
            
            return outTable;
        }

        /// <summary>
        /// Will return a table object from a full path reference (.shp and FCs)
        /// </summary>
        /// <param name="inputFeature"> Table full reference path</param>
        /// <returns></returns>
        public static ITable OpenTableFromString(string inputTablePath)
        {
            //Call the GPUtilities function to retrieve feature class (treats shapefiles and features)
            IGPUtilities gpUtil = new GPUtilities();
            ITable getInTable = gpUtil.OpenTableFromString(inputTablePath);
            Services.ObjectManagement.ReleaseObject(gpUtil);
            return getInTable;
        }

        /// <summary>
        /// Will return a table object from a full path reference (.shp and FCs)
        /// </summary>
        /// <param name="inputFeature"> Table full reference path</param>
        /// <returns></returns>
        public static ITable OpenTableFromStringFaster(string inputTablePath)
        {
            //Get some information from full path
            string dbPath = inputTablePath;// System.IO.Directory.GetParent(inputTablePath).FullName;
            string[] splitedPath = inputTablePath.Split('\\');
            string fileName = splitedPath[splitedPath.Length - 1];

            //Validate if feature class isn't inside a feature dataset,else redo dbPath
            if (dbPath.Contains(".gdb") || dbPath.Contains(".mdb"))
            {
                //Get parent folder
                dbPath = System.IO.Directory.GetParent(inputTablePath).FullName;

                if (dbPath.Substring(dbPath.Length - 4, 4).ToLower() != ".gdb" && dbPath.Substring(dbPath.Length - 4, 4).ToLower() != ".mdb")
                {
                    //Reset path to nothing
                    dbPath = string.Empty;

                    //Cast to list
                    List<string> splittedPathList = new List<string>(splitedPath);
                    splittedPathList.RemoveRange(splittedPathList.Count - 2, 2);

                    foreach (string parts in splittedPathList)
                    {
                        if (dbPath != string.Empty)
                        {
                            dbPath = dbPath + "\\" + parts;
                        }
                        else
                        {
                            dbPath = parts;
                        }
                    }
                }
            }
            else
            {
            }

            //Access the workspace
            IWorkspace inputWorkspace = Services.Workspace.AccessWorkspace(dbPath);

            //Access the table behind the dataset
            ITable inputDatasetTable = Services.Tables.OpenTableFromWorkspace(inputWorkspace, fileName);

            Services.ObjectManagement.ReleaseObject(inputWorkspace);

            return inputDatasetTable;
        }

        /// <summary>
        /// Will return a table object for first item found with table name inside name of item.
        /// </summary>
        /// <param name="inputWorkspace">The workspace to find the table name in</param>
        /// <param name="inputTableName">The table name to search for in all table names</param>
        /// <returns></returns>
        public static ITable OpenTableFromWorkspaceWildCard(IWorkspace inputWorkspace, string inputTableName)
        {
            //Variables
            ITable foundTable = default(ITable);

            //Get a list of all tables
            List<string> tableList = GetTableNameListFromWorkspace(inputWorkspace);

            //Find the wanted table name from list
            foreach (string tableNames in tableList)
            {
                if (tableNames.Contains(inputTableName))
                {
                    //open
                    foundTable = OpenTableFromWorkspace(inputWorkspace, tableNames);
                }
            }

            return foundTable;

        }

        #endregion

        #region ADD methods
        /// <summary>
        /// Add A new row into table, and fill field values, from any given workspace
        /// </summary>
        /// <param name="inputWorkspace">The workspace to add a row into table</param>
        /// <param name="inTableName">The table to update</param>
        /// <param name="inFieldAndValues">The fields and values contained in a dictionnary</param>
        public static void AddRowWithValuesFromWorkspace(IWorkspace inputWorkspace, string inTableName, Dictionary<string, object> inFieldAndValues)
        {
            List<Dictionary<string, object>> inFieldAndValuesList = new List<Dictionary<string, object>> { inFieldAndValues };
            AddRowsWithValuesFromWorkspace(inputWorkspace, inTableName, inFieldAndValuesList);
        }

        /// <summary>
        /// Add new row into table, and fill field values, from any given workspace
        /// </summary>
        /// <param name="inputWorkspace">The workspace to add a row into table</param>
        /// <param name="inTableName">The table to update</param>
        /// <param name="inFieldAndValues">The fields and values contained in a dictionnary</param>
        public static void AddRowsWithValuesFromWorkspace(IWorkspace inputWorkspace, string inTableName, List<Dictionary<string, object>> inFieldAndValues)
        {
            //Get table object
            ITable inWantedTable = OpenTableFromWorkspace(inputWorkspace, inTableName);

            //Start a cursor to insert new row
            ICursor inCursor = inWantedTable.Insert(true);

            foreach(Dictionary<string, object> dico in inFieldAndValues)
            {
                //Create a row buffer object (a template of all fields of table)
                IRowBuffer inRowBuffer = inWantedTable.CreateRowBuffer();

                //Iterate through dictionnary
                foreach (KeyValuePair<string, object> elements in dico)
                {
                    //Get field index
                    int fieldIndex = inRowBuffer.Fields.FindField(elements.Key);

                    //Set field value
                    inRowBuffer.set_Value(fieldIndex, elements.Value);
                }

                //Add the new row
                inCursor.InsertRow(inRowBuffer);
                
            }
            inCursor.Flush();
            //Release the cursor or else some lock could happen.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(inCursor);
        }

        /// <summary>
        /// Will add a new field to in input project database feature
        /// </summary>
        /// <param name="fieldParam">A tuple containing in this order the params of the new field (name, field type, isNullable, AliasName, default value, editable, length)</param>
        public static void AddField(Tuple<string, esriFieldType, bool, string, object, bool, object> fieldParam, ITable inputTable)
        {
            try
            {
                //Create the field object
                IField newField = new Field();
                IFieldEdit2 newFieldEdit = (IFieldEdit2)newField;

                newFieldEdit.Name_2 = fieldParam.Item1;
                newFieldEdit.Type_2 = fieldParam.Item2;
                newFieldEdit.IsNullable_2 = fieldParam.Item3;
                newFieldEdit.AliasName_2 = fieldParam.Item4;
                newFieldEdit.DefaultValue_2 = fieldParam.Item5;
                newFieldEdit.Editable_2 = fieldParam.Item6;
                if (fieldParam.Item7 != null)
                {
                    newFieldEdit.Length_2 = Convert.ToInt16(fieldParam.Item7);
                }

                //Add to feature
                inputTable.AddField(newField);
            }
            catch (Exception addFieldExcept)
            {
                int lineNumber = Services.Exceptions.LineNumber(addFieldExcept);
                MessageBox.Show("addFieldExcept (" + lineNumber.ToString() + "): " + addFieldExcept.Message);
            }

        }

        #endregion

        #region UPDATE methods

        /// <summary>
        /// Update a table field with wanted values
        /// </summary>
        /// <param name="upTableName">Table name to update</param>
        /// <param name="upFieldName">Field name to update</param>
        /// <param name="upQuery">Query if needed, else reference a null string</param>
        /// <param name="upValue">Field value that will be update</param>
        public static void UpdateFieldValueFromWorkspace(IWorkspace inputWorkspace, string upTableName, string upFieldName, string upQuery, object upValue)
        {
            try
            {
                //Variables
                IQueryFilter queryFilterUpdate = new QueryFilter();

                //Get table object
                ITable upWantedTable = OpenTableFromWorkspace(inputWorkspace, upTableName);

                //Get field index
                int upFieldIndex = upWantedTable.FindField(upFieldName);

                //Get a search cursor within table
                if (upQuery != null)
                {
                    queryFilterUpdate.WhereClause = upQuery;
                }
                else
                {
                    queryFilterUpdate = null;
                }

                ICursor upCursor = upWantedTable.Update(queryFilterUpdate, true);
                IRow upRows = null;
                while ((upRows = upCursor.NextRow()) != null)
                {
                    upRows.set_Value(upFieldIndex, upValue);
                    upCursor.UpdateRow(upRows);
                    upCursor.Flush();
                }

                //Release the cursor or else some lock could happen.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(upCursor);

            }
            catch (Exception UpdateFieldValueFromWorkspaceError)
            {
                MessageBox.Show("UpdateFieldValueFromWorkspace: " + UpdateFieldValueFromWorkspaceError.Message);
            }
        }

        /// <summary>
        /// Update a table multiple fields with wanted values, from an input dictionnary and input workspace
        /// </summary>
        /// <param name="inputWorkspace">Input workspace object</param>
        /// <param name="updateTableName">Table name to update</param>
        /// <param name="updateQuery">Query if needed, else reference a null string</param>
        /// <param name="updateDictionnary">Dictionnary containing as keys fields names, and values well field values</param>
        public static void UpdateMultipleFieldValueFromWorkspace(IWorkspace inputWorkspace, string updateTableName, string updateQuery, Dictionary<string, object> updateDictionnary)
        {
            try
            {
                ICursor updateCursor = GetTableCursorFromWorkspace(inputWorkspace, "Update", updateQuery, updateTableName);
                IRow updateRows = null;
                while ((updateRows = updateCursor.NextRow()) != null)
                {
                    foreach (KeyValuePair<string, object> updateKeyValues in updateDictionnary)
                    {
                        //Get field index
                        int currentFielIndex = updateRows.Fields.FindField(updateKeyValues.Key);

                        //Update value 
                        updateRows.set_Value(currentFielIndex, updateKeyValues.Value);
                    }

                    //Persist changes
                    updateCursor.UpdateRow(updateRows);
                }
                updateCursor.Flush();

                //Release the cursor or else some lock could happen.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(updateCursor);

            }
            catch (Exception UpdateMultipleFieldValueError)
            {
                MessageBox.Show("UpdateMultipleFieldValueError: " + UpdateMultipleFieldValueError.Message + Services.Exceptions.LineNumber(UpdateMultipleFieldValueError));
            }
        }

        #endregion

        #region DELETE methods

        /// <summary>
        /// Delete a row from a table.
        /// </summary>
        /// <param name="inputWorkspace">The given workspace to get the table to delete from</param>
        /// <param name="delTableName"> The table name to delete a row from</param>
        /// <param name="delQuery"> The query to select proper values to delete</param>
        public static void DeleteFieldValueFromWorkspace(IWorkspace inputWorkspace, string delTableName, string delQuery)
        {
            try
            {
                //Variables
                IQueryFilter queryFilterUpdate = new QueryFilter();

                //Get table object
                ITable delWantedTable = OpenTableFromWorkspace(inputWorkspace, delTableName);

                //Get a update cursor within table, need to have a query, in case something bad happens, the table won't be erase in it's enterity
                if (delQuery != null)
                {
                    queryFilterUpdate.WhereClause = delQuery;
                    ICursor upCursor = delWantedTable.Update(queryFilterUpdate, true);
                    IRow upRows = upCursor.NextRow();
                    while (upRows != null)
                    {
                        upRows.Delete();
                        upRows = upCursor.NextRow();
                    }
                    //Release the cursor or else some lock could happen.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(upCursor);
                }

            }
            catch (Exception DeleteFieldValueFromWorkspaceError)
            {
                MessageBox.Show("DeleteFieldValueFromWorkspaceError: " + DeleteFieldValueFromWorkspaceError.Message);
            }
        }

        /// <summary>
        /// Will delete a given table from it's workspace
        /// </summary>
        /// <param name="tableToDelete">The table to delet</param>
        public static void DeleteTableFromWorkspace(ITable tableToDelete)
        {
            //Cast as dataset
            IDataset tableDataset = tableToDelete as IDataset;

            //Delete
            tableDataset.Delete();
        }

        /// <summary>
        /// Will delete from given table, a given field.
        /// </summary>
        /// <param name="inTable">The table in which a field will be deleted</param>
        /// <param name="delField">The field to delete</param>
        public static void DeleteField(ITable inTable, string fieldToDelete)
        {
            //Get table object
            IFields currentFields = inTable.Fields;
            for (int fieldCount = 0; fieldCount <= currentFields.FieldCount; fieldCount++)
            {
                if (currentFields.Field[fieldCount].Name == fieldToDelete)
                {
                    inTable.DeleteField(currentFields.Field[fieldCount]);
                }
            }

        }

        public static void EmptyTable(IWorkspace inputWorkspace, string delTableName)
        {
            try
            {
                //Variables
                IQueryFilter queryFilterUpdate = new QueryFilter();

                //Get table object and delete
                ITable delWantedTable = OpenTableFromWorkspace(inputWorkspace, delTableName);
                delWantedTable.DeleteSearchedRows(queryFilterUpdate);

            }
            catch (Exception emptyTableException)
            {
                MessageBox.Show("emptyTableException: " + emptyTableException.StackTrace);
            }
        }

        #endregion

        #region COPY methods

        /// <summary>
        /// Will copy a given table to project database with a new given name
        /// </summary>
        /// <param name="inTable">Table to copy</param>
        /// <param name="outName">New name</param>
        public static void CopyTableToWorkspace(IWorkspace outWorkspace, ITable inTable, string outName)
        {
            //Cast input table as dataset and access it's name
            IDataset inDataset = inTable as IDataset;
            IDatasetName inDatasetName = inDataset.FullName as IDatasetName;

            //Init a new table name and set the new one
            ITableName outTableName = new TableName() as ITableName;
            IDatasetName outDatasetName = outTableName as IDatasetName;
            outDatasetName.Name = outName;

            //Set workspace of output table
            IDataset outDataset = outWorkspace as IDataset;
            IWorkspaceName outWorkspaceName = new WorkspaceName() as IWorkspaceName;
            outWorkspaceName = outDataset.FullName as IWorkspaceName;
            outDatasetName.WorkspaceName = outWorkspaceName;

            //Export
            IExportOperation newExport = new ExportOperation();
            newExport.ExportTable(inDatasetName, null, null, outDatasetName, 0);
            
        }

        #endregion
    }
}
