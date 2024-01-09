using System;
using System.Collections.Generic;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using SimioAPI;
using SimioAPI.Extensions;

//using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using OfficeOpenXml;
using System.Linq;

namespace ImportObjectsFromExcelUsingEPPlus
{
    public class ImportObjectsFromExcelUsingEPPlus : IDesignAddIn
    {
        #region IDesignAddIn Members

        /// <summary>
        /// Property returning the name of the add-in. This name may contain any characters and is used as the display name for the add-in in the UI.
        /// </summary>
        public string Name
        {
            get { return "Load objects, links, and vertices from Excel spreadsheet using EPPlus"; }
        }

        /// <summary>
        /// Property returning a short description of what the add-in does.  
        /// </summary>
        public string Description
        {
            get { return "This add-in loads objects, links, and vertices into a model from an Excel spreadsheet using the OpenSource EPPlus library (from GitHub)"; }
        }

        /// <summary>
        /// Property returning an icon to display for the add-in in the UI.
        /// </summary>
        public Image Icon
        {
            get { return Properties.Resources.Icon; }
        }

        #endregion

        private StatusWindow StatusWindow = null;

        /// <summary>
        /// Method called when the add-in is run.
        /// </summary>
        public void Execute(IDesignContext context)
        {
            //Open Status Window
            string marker = "Begin.";

            try
            {

                // Check to make sure a model has been opened in Simio
                if (context.ActiveModel == null)
                {
                    MessageBox.Show("You must have an active model to run this add-in.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Open the file.  Return immediately if the user cancels the file open dialog
                var getFile = new OpenFileDialog();
                getFile.Filter = "Excel Files(*.xlsx)|*.xlsx";
                if (getFile.ShowDialog() == DialogResult.Cancel)
                {
                    MessageBox.Show("Canceled by user.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                StatusWindow = new StatusWindow($"Importing From Excel Spreadsheet {getFile.FileName}");
                StatusWindow.Show();

                Boolean importVertices = true;

                // Update Status Window
                StatusWindow.UpdateProgress(25, "Checking worksheets");

                ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(getFile.FileName));
                ExcelWorkbook xlWorkbook = package.Workbook;

                // Create the node and link sheet lists
                List<ExcelWorksheet> objectsWorksheets = new List<ExcelWorksheet>();
                List<ExcelWorksheet> linksWorksheets = new List<ExcelWorksheet>();
                List<ExcelWorksheet> verticesWorksheets = new List<ExcelWorksheet>();

                marker = "Categorizing Worksheets.";
                // Look through every sheet and categorize them according to objects, links, or vertices
                // We'll do objects first, then vertices, then links
                foreach (ExcelWorksheet ws in package.Workbook.Worksheets)
                {
                    string wsName = ws.Name.ToLower();

                    if (wsName.Length >= 5)
                    {
                        // Add any sheet that name starts with 'objects' to the objects list
                        if (wsName.ToLower().StartsWith("objects"))
                        {
                            objectsWorksheets.Add(ws);
                        }
                        // Add any sheet that name starts with 'links' to the link list
                        else if (wsName.ToLower().StartsWith("links"))
                        {
                            linksWorksheets.Add(ws);
                        }
                        // Add any sheet that name starts with 'vertices' to the link list
                        else if (wsName.ToLower().StartsWith("vertices"))
                        {
                            verticesWorksheets.Add(ws);
                        }
                    }
                } // foreach worksheet

                if (objectsWorksheets.Count + linksWorksheets.Count == 0)
                {
                    logit("Workbook contains no valid object or link worksheets.");
                    return;
                }

                // Update Status Window
                StatusWindow.UpdateProgress(50, "Building Objects");

                // get a reference to intelligent objects...
                var intellObjects = context.ActiveModel.Facility.IntelligentObjects;

                // ... and a reference to Elements
                var elements = context.ActiveModel.Elements;

                // use bulk update to import quicker
                context.ActiveModel.BulkUpdate(model =>
                {
                    // Read and create the objects.
                    int addedCount;
                    int updatedCount;
                    foreach (ExcelWorksheet ows in objectsWorksheets)
                    {
                        var dim = ows.Dimension;
                        if (dim.Rows == 0)
                            continue;

                        marker = $"Reading {dim.Rows} rows from Object sheet {ows.Name}";
                        logit(marker);

                        addedCount = 0;
                        updatedCount = 0;

                        for (int ri = 2; ri <= dim.Rows; ri++)
                        {
                            marker = $"Sheet={ows.Name} Row={ri}";

                            string className = ows.Cells[ri, 1].Value?.ToString();
                            string itemName = ows.Cells[ri, 2].Value?.ToString();

                            if (string.IsNullOrEmpty(className) || string.IsNullOrEmpty(itemName))
                            {
                                logit($"{marker}: Empty ClassName or ItemName");
                                continue;
                            }

                            // Find the coordinates for the object
                            double x = 0.0, y = 0.0, z = 0.0;
                            bool updateCoordinates = true;

                            if (!GetCellAsDouble(ows.Cells[ri, 3], ref x)
                                || !GetCellAsDouble(ows.Cells[ri, 4], ref y)
                                || !GetCellAsDouble(ows.Cells[ri, 5], ref z))
                                updateCoordinates = false;

                            // Add the coordinates to the intelligent object
                            FacilityLocation loc = new FacilityLocation(x, y, z);
                            var intellObj = intellObjects[itemName];
                            if (intellObj == null)
                            {
                                intellObj = intellObjects.CreateObject(className, loc);
                                if (intellObj == null)
                                {
                                    logit($"{marker}: Cannot create object with className={className}");
                                    continue;
                                }

                                intellObj.ObjectName = itemName;
                                addedCount++;
                            }
                            else
                            {
                                // update coords of existing one.
                                if (updateCoordinates)
                                {
                                    intellObj.Location = loc;
                                }
                                updatedCount++;
                            }

                            // Set Size
                            double length = intellObj.Size.Length;
                            double width = intellObj.Size.Width;
                            double height = intellObj.Size.Height;

                            if (GetCellAsDouble(ows.Cells[ri, 6], ref length))
                                if (length == 0) length = intellObj.Size.Length;

                            if (GetCellAsDouble(ows.Cells[ri, 7], ref width))
                                if (width == 0) width = intellObj.Size.Width;

                            if (GetCellAsDouble(ows.Cells[ri, 8], ref height))
                                if (height == 0) height = intellObj.Size.Height;

                            FacilitySize fs = new FacilitySize(length, width, height);
                            intellObj.Size = fs;


                            // update properties on object, which are columns 9 onward
                            for (int ci = 9; ci <= dim.Columns; ci++)
                            {
                                // By convention, the first row on the sheet is the header row, which contains the Property name.
                                string propertyName = ows.Cells[1, ci]?.Value as string;
                                if (string.IsNullOrEmpty(propertyName))
                                    continue;

                                propertyName = propertyName.ToLower();
                                
                                // Find a property with matching text (case insensitive)
                                IProperty prop = intellObj.Properties.AsQueryable()
                                        .SingleOrDefault(rr => rr.Name.ToString().ToLower() == propertyName);

                                if (prop == null)
                                    continue;

                                string cellValue = ows.Cells[ri, ci].Value?.ToString();
                                if (cellValue != null)
                                {
                                    if (!SetPropertyValue(prop, cellValue, out string explanation))
                                    {
                                        logit(explanation);
                                    }
                                }
                            } // for each column property
                        } // foreach row
                        logit($"Added {addedCount} objects and updated {updatedCount} objects");
                    } // for each object worksheet

                    var vertexList = new ArrayList();

                    // Update Status Window
                    if (importVertices)
                    {
                        //  Add additional vertices
                        foreach (ExcelWorksheet vws in verticesWorksheets)
                        {
                            var dim = vws.Dimension;
                            if (dim.Rows > 0)
                            {
                                logit($"Info: Reading {dim.Rows} rows from sheet {vws.Name}");
                            }
                            addedCount = 0;
                            updatedCount = 0;

                            for (int ri = 2; ri <= dim.Rows; ri++)
                            {
                                marker = $"Sheet={vws.Name} Row={ri}";

                                var cell = vws.Cells[ri, 1];
                                string linkName = cell.Value as string;
                                if (string.IsNullOrEmpty(linkName))
                                {
                                    logit($"{marker}: No LinkName");
                                    goto DoneWithVertexRows;
                                }
                                // Find the coordinates for the vertex
                                double x = double.MinValue, y = double.MinValue, z = double.MinValue;
                                if (    !GetCellAsDouble(vws.Cells[ri, 2], ref x)
                                    ||  !GetCellAsDouble(vws.Cells[ri, 3], ref y)
                                    ||  !GetCellAsDouble(vws.Cells[ri, 4], ref z))
                                {
                                    logit($"{marker}: Bad Vertex Coordinate");
                                    goto DoneWithVertexRows;
                                }

                                vertexList.Add(new string[] { linkName, x.ToString(), y.ToString(), z.ToString() });
                            } // for each row of vertices

                            DoneWithVertexRows:;
                        } // for each vertex worksheet
                    } // Check if we are importing vertices

                    StatusWindow.UpdateProgress(75, "Building Links");

                    // Get Links Data

                    // Read and create the links.
                    foreach (ExcelWorksheet lws in linksWorksheets)
                    {
                        var dim = lws.Dimension;
                        if (dim.Rows > 0)
                        {
                            marker = $"Info: Reading {dim.Rows} rows from sheet {lws.Name}";
                            logit(marker);
                        }
                        addedCount = 0;
                        updatedCount = 0;

                        for (int ri = 2; ri <= dim.Rows; ri++)
                        {
                            marker = $"Sheet={lws.Name} Row={ri}";
                            string className = lws.Cells[ri, 1]?.Value as string;
                            if (string.IsNullOrEmpty(className))
                            {
                                logit($"{marker}: Invalid ClassName={className}");
                                goto DoneWithLinkRow;
                            }

                            string linkName = lws.Cells[ri, 2]?.Value as string;
                            if (string.IsNullOrEmpty(linkName))
                            {
                                logit($"{marker}: Invalid LinkName={linkName}");
                                goto DoneWithLinkRow;
                            }

                            string fromNodeName = lws.Cells[ri, 3]?.Value as string;
                            if (string.IsNullOrEmpty(fromNodeName))
                            {
                                logit($"{marker}: Invalid FromNodeName={fromNodeName}");
                                goto DoneWithLinkRow;
                            }
                            string toNodeName = lws.Cells[ri, 4]?.Value as string;
                            if (string.IsNullOrEmpty(toNodeName))
                            {
                                logit($"{marker}: Invalid ToNodeName={toNodeName}");
                                goto DoneWithLinkRow;
                            }


                            var fromNode = intellObjects[fromNodeName] as INodeObject;
                            if (fromNode == null)
                            {
                                logit($"{marker} Cannot find 'from' node name {fromNodeName}");
                                goto DoneWithWorksheets;
                            }

                            var toNode = intellObjects[toNodeName] as INodeObject;
                            if (toNode == null)
                            {
                                logit($"{marker}: Cannot find 'to' node name {toNodeName}");
                                goto DoneWithWorksheets;
                            }

                            // if link exists, remove and re-add
                            var link = intellObjects[linkName];
                            if (link != null)
                            {
                                intellObjects.Remove(link);
                                updatedCount++;
                            }
                            else addedCount++;

                            // Define List of Facility Locations
                            List<FacilityLocation> locs = new List<FacilityLocation>();

                            foreach (string[] loc in vertexList)
                            {
                                if (loc[0] == linkName)
                                {
                                    double xx = double.MinValue, yy = double.MinValue, zz = double.MinValue;

                                    xx = Convert.ToDouble(loc[1]);
                                    yy = Convert.ToDouble(loc[2]);
                                    zz = Convert.ToDouble(loc[3]);

                                    // If coordinates are good, add vertex to facility locations
                                    if (xx > double.MinValue & yy > double.MinValue & zz > double.MinValue)
                                    {
                                        // Add the coordinates to the intelligent object
                                        FacilityLocation loc2 = new FacilityLocation(xx, yy, zz);
                                        locs.Add(loc2);
                                    }
                                }
                            } // for each vertex

                            // Add Link
                            link = intellObjects.CreateLink(className, fromNode, toNode, locs);
                            if (link == null)
                            {
                                logit($"{marker}: Cannot create Link");
                                goto DoneWithWorksheets;
                            }
                            link.ObjectName = linkName;

                            // Add Link to Network
                            string networkName = lws.Cells[ri, 5]?.Value as string;
                            if (string.IsNullOrEmpty(networkName))
                            {
                                logit($"{marker}: Null NetworkName");
                                goto DoneWithLinkRow;
                            }

                            var netElement = elements[networkName];
                            if (netElement == null)
                            {
                                netElement = elements.CreateElement("Network");
                                netElement.ObjectName = networkName;
                            }
                            var netElementObj = netElement as INetworkElementObject;
                            if (netElement != null)
                            {
                                ILinkObject linkOb = (ILinkObject)link;
                                linkOb.Networks.Add(netElementObj);
                            }

                            // get header row on sheet

                            // update properties on object, which begin with column index 6
                            for (int ci = 6; ci <= dim.Columns; ci++)
                            {
                                string propertyName = lws.Cells[1, ci]?.Value as string;

                                // Find a property with matching text (case insensitive)
                                IProperty prop = link.Properties.AsQueryable()
                                        .SingleOrDefault(rr => rr.Name.ToString().ToLower() == propertyName.ToLower());

                                if (prop != null)
                                {
                                    string cellValue = lws.Cells[ri, ci]?.Value as string;
                                    if (!SetPropertyValue(prop, cellValue, out string explanation))
                                    {
                                        logit(explanation);
                                        goto DoneWithLinkRow;
                                    }
                                }

                            } // foreach column

                            DoneWithLinkRow:;
                        }   // for each row
                        marker =$"Info: Added {addedCount} links and deleted and re-added {updatedCount} existing links";
                        logit(marker);

                        DoneWithWorksheets:;
                    }
                });

                // Update Status Window
                StatusWindow.UpdateProgress(100, "Complete");

            }
            catch (Exception ex)
            {

                throw new ApplicationException($"Marker={marker} Err={ex.Message}");
            }
            finally
            {
                //StatusWindow.UpdateLogs(Loggerton.Instance);
            }

        }

        /// <summary>
        /// Ok, we have a property. Is it a repeating property?
        /// If so, the value has a special format. An example is:
        /// "1 Row;AssignmentsOnEnteringStateVariableName;ModelEntity.Picture~AssignmentsOnEnteringNewValue;1"
        /// There are tokens split with a tilde (~). The first one is "1 Row"
        /// The others are property pairs, separated by a semicolon (;)
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="value"></param>
        /// <param name="explanation"></param>
        /// <returns></returns>
        public bool SetPropertyValue(IProperty prop, string cellValue, out string explanation)
        {
            explanation = "";

            try
            {
                if (prop is IRepeatingProperty)
                {
                    var repeatProp = prop as IRepeatingProperty;

                    // repeating property 
                    Int32 prevRowNumber = -1;
                    string[] tokens1 = cellValue.Split('~');
                    for (int ii = 0; ii < tokens1.Length; ii++)
                    {
                        string[] tokens2 = tokens1[ii].Split(';');

                        if (ii == 0 && tokens2.Length > 2)
                        {
                            prop.Value = tokens2[0];
                        }

                        if (tokens2.Length >= 2)
                        {
                            Int32 currentRowNumber = -1;
                            bool result = false;

                            if (tokens2.Length > 2)
                                result = Int32.TryParse(tokens2[tokens2.Length - 3], out currentRowNumber);

                            IRow propRow;
                            if (result == true && prevRowNumber < currentRowNumber)
                            {
                                prevRowNumber = currentRowNumber;
                                if (currentRowNumber > repeatProp.Rows.Count)
                                    propRow = repeatProp.Rows.Create();
                                else
                                    propRow = repeatProp.Rows[currentRowNumber - 1];

                            }
                            else
                            {
                                if (repeatProp.Rows.Count == 0)
                                    propRow = repeatProp.Rows.Create();
                                else
                                    propRow = repeatProp.Rows[repeatProp.Rows.Count - 1];
                            }

                            foreach (var propRowProp in propRow.Properties)
                            {
                                if (tokens2[tokens2.Length - 2] == propRowProp.Name)
                                {
                                    propRowProp.Value = tokens2[tokens2.Length - 1];
                                    break;
                                }
                            } // foreach property in 
                        }
                    }

                }
                else
                {
                    prop.Value = cellValue;
                }

                return true;
            }
            catch (Exception ex)
            {
                explanation = $"Property={prop.Name} Value={cellValue} Err={ex}";
                return false;
            }
        }

        /// <summary>
        /// Given a range, get a list of doubles
        /// </summary>
        /// <param name="range"></param>
        /// <param name="ddList"></param>
        /// <returns></returns>
        public bool GetRangeAsDoubleList(ExcelRange range, out List<double> ddList)
        {
            ddList = new List<double>();
            string marker = "Begin";
            try
            {
                object[,] arr = range.Value as object[,];
                for (int rr = 0; rr < arr.GetLength(0); rr++)
                {
                    for (int cc = 0; cc < arr.GetLength(1); rr++)
                    {
                        marker = $"[{rr},{cc}]";
                        double dd = 0D;

                        var val = (string)arr[rr, cc];
                        if (val == null || !double.TryParse((string)val, out dd))
                        {
                            return false;
                        }
                        else
                            ddList.Add(dd);
                    } // for each column
                } // for each row
                return true;
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Marker={marker} Err={ex.Message}");
            }
        }

        /// <summary>
        /// Get a cell as text and try and parse it as a double.
        /// Return false if the value is null or isn't a decimal, in which case the dd argument is untouched.
        /// Return true if a legitimate double (dd) is found and set.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dd"></param>
        /// <returns></returns>
        public bool GetCellAsDouble(ExcelRange cell, ref double dd)
        {
            if (cell?.Value == null)
                return false;

            double newValue;

            if (Double.TryParse(cell.Text, out newValue))
            {
                dd = newValue;
                return true;
            }
            else
                return false;

        }

        /// <summary>
        /// Log message to the internal loggerton instance.
        /// </summary>
        /// <param name="message"></param>
        private void logit(string message)
        {
            if (StatusWindow == null)
                return;

            Loggerton.Instance.LogIt(message);
        }
    } // class
} // namespace
