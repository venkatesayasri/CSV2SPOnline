using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using CSV2SPOnline.Properties;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using System.Security;

namespace CSV2SPOnline
{
    class Program
    {
        static SecureString GetPasswordFromConsole()
        {
            SecureString _secureString = new SecureString();

            try
            {
                for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
                {
                    if (keyInfo.Key == ConsoleKey.Backspace)
                    {
                        if (_secureString.Length > 0)
                        {
                            _secureString.RemoveAt(_secureString.Length - 1);
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                            Console.Write(" ");
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        }
                    }
                    else if (keyInfo.Key != ConsoleKey.Enter)
                    {
                        Console.Write("*");
                        _secureString.AppendChar(keyInfo.KeyChar);
                    }

                }
                Console.WriteLine("");
            }
            catch (Exception e)
            {
                _secureString = null;
                Console.WriteLine(e.Message);
            }

            return _secureString;
        }

        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(Settings.Default.FilePath, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                foreach (Cell cell in rows.ElementAt(0))
                {
                    dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }
                foreach (Row row in rows) //this will also include your header row...
                {
                    DataRow tempRow = dt.NewRow();
                    int columnIndex = 0;
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        // Gets the column index of the cell with data
                        int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
                        cellColumnIndex--; //zero based index
                        if (columnIndex < cellColumnIndex)
                        {
                            do
                            {
                                tempRow[columnIndex] = ""; //Insert blank data here;
                                columnIndex++;
                            }
                            while (columnIndex < cellColumnIndex);
                        }
                        tempRow[columnIndex] = GetCellValue(spreadSheetDocument, cell);

                        columnIndex++;
                    }
                    dt.Rows.Add(tempRow);
                }
            }
            dt.Rows.RemoveAt(0);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your user name (ex: employeeID@croydon.gov.uk):");
            string _targetUserName = Console.ReadLine();


            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your password:");
            SecureString _targetPassword = GetPasswordFromConsole();

            ClientContext ctx = new ClientContext(Settings.Default.SiteURL);
            ctx.Credentials = new SharePointOnlineCredentials(_targetUserName, _targetPassword);
            ctx.ExecuteQuery();

            List targetList = ctx.Web.Lists.GetByTitle(Settings.Default.ListName);
            ctx.Load(targetList);
            ctx.ExecuteQuery();

            foreach (DataRow dr in dt.Rows)
            {
                addNewRecord(dr, targetList, ctx);
            }
        }


        public static void addNewRecord(DataRow dr, List targetList, ClientContext ctx)
        {
            //Instantiate dictionary to temporarily store field values
            Dictionary<string, object> itemFieldValues = new Dictionary<string, object>();
            foreach (DataColumn dc in dr.Table.Columns)
            {
                //Get site column that matches the property name
                //ASSUMPTION: Your property names match the internal names of the corresponding site columns
                Microsoft.SharePoint.Client.Field matchingField = targetList.Fields.GetByInternalNameOrTitle(dc.ColumnName);
                ctx.Load(matchingField);
                ctx.ExecuteQuery();

                //Switch on the field type
                switch (matchingField.FieldTypeKind)
                {
                    case FieldType.DateTime:
                        try
                        {
                            DateTime date = DateTime.Parse(dr[dc].ToString());
                            itemFieldValues.Add(matchingField.InternalName, date);
                        }
                        catch(Exception ex)
                        {

                        }
                        break;
                    case FieldType.User:
                        FieldUserValue userFieldValue = GetUserFieldValue(dr[dc].ToString(), ctx);
                        if (userFieldValue != null)
                            itemFieldValues.Add(matchingField.InternalName, userFieldValue);
                        else
                            throw new Exception("User field value could not be added: " + dr[dc].ToString());
                        break;
                    case FieldType.Lookup:
                        var lookupField = ctx.CastTo<FieldLookup>(matchingField);
                         ctx.Load(lookupField);
                         ctx.ExecuteQuery();
                        FieldLookupValue lookupFieldValue = GetLookupFieldValue(dr[dc].ToString(),
                            lookupField.LookupList, lookupField.LookupField,
                            ctx);
                        if (lookupFieldValue != null)
                            itemFieldValues.Add(matchingField.InternalName, lookupFieldValue);
                        else
                            throw new Exception("Lookup field value could not be added: " + dr[dc].ToString());
                        break;
                    case FieldType.Invalid:
                        switch (matchingField.TypeAsString)
                        {
                            case "TaxonomyFieldType":
                                TaxonomyFieldValue taxFieldValue = GetTaxonomyFieldValue(dr[dc].ToString(), matchingField, ctx);
                                if (taxFieldValue != null)
                                    itemFieldValues.Add(matchingField.InternalName, taxFieldValue);
                                else
                                    throw new Exception("Taxonomy field value could not be added: " + dr[dc].ToString());
                                break;
                            default:
                                //Code for publishing site columns not implemented
                                continue;
                        }
                        break;
                    default:
                        itemFieldValues.Add(matchingField.InternalName, dr[dc]);
                        break;
                }
            }

            //Add new item to list
            ListItemCreationInformation creationInfo = new ListItemCreationInformation();
            ListItem oListItem = targetList.AddItem(creationInfo);

            foreach (KeyValuePair<string, object> itemFieldValue in itemFieldValues)
            {
                //Set each field value
                oListItem[itemFieldValue.Key] = itemFieldValue.Value;
            }
            //Persist changes
            oListItem.Update();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(oListItem["Title"].ToString() + " : Inserted!!");
            ctx.ExecuteQuery();
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Column Name (ie. B)</returns>
        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }
        /// <summary>
        /// Given just the column name (no row index), it will return the zero based column index.
        /// Note: This method will only handle columns with a length of up to two (ie. A to Z and AA to ZZ). 
        /// A length of three can be implemented when needed.
        /// </summary>
        /// <param name="columnName">Column Name (ie. A or AB)</param>
        /// <returns>Zero based index if the conversion was successful; otherwise null</returns>
        public static int? GetColumnIndexFromName(string columnName)
        {

            //return columnIndex;
            string name = columnName;
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            return number;
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue == null)
            {
                return "";
            }
            string value = cell.CellValue.InnerXml;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText.Replace("_x005F_", "_");
            }
            else
            {
                return value;
            }
        }

        private static FieldUserValue GetUserFieldValue(string userName, ClientContext clientContext)
        {
            //Returns first principal match based on user identifier (display name, email, etc.)
            ClientResult<PrincipalInfo> principalInfo = Utility.ResolvePrincipal(
                clientContext, //context
                clientContext.Web, //web
                userName, //input
                PrincipalType.User, //scopes
                PrincipalSource.All, //sources
                null, //usersContainer
                false); //inputIsEmailOnly
            clientContext.ExecuteQuery();
            PrincipalInfo person = principalInfo.Value;

            if (person != null)
            {
                //Get User field from login name
                User validatedUser = clientContext.Web.EnsureUser(person.LoginName);
                clientContext.Load(validatedUser);
                clientContext.ExecuteQuery();

                if (validatedUser != null && validatedUser.Id > 0)
                {
                    //Sets lookup ID for user field to the appropriate user ID
                    FieldUserValue userFieldValue = new FieldUserValue();
                    userFieldValue.LookupId = validatedUser.Id;
                    return userFieldValue;
                }
            }
            return null;
        }

        public static FieldLookupValue GetLookupFieldValue(string lookupName, string lookupListName, string lookupFieldName, ClientContext clientContext)
        {
            var lookupList = clientContext.Web.Lists.GetById(new Guid(lookupListName));
            CamlQuery query = new CamlQuery();

            query.ViewXml = string.Format(@"<View><Query><Where><Eq><FieldRef Name='{0}'/><Value Type='Text'>{1}</Value></Eq>" +
                                            "</Where></Query></View>", lookupFieldName, lookupName);

            ListItemCollection listItems = lookupList.GetItems(query);
            clientContext.Load(listItems, items => items.Include
                                                (listItem => listItem["ID"],
                                                listItem => listItem[lookupFieldName]));
            clientContext.ExecuteQuery();

            if (listItems != null && listItems.Count > 0)
            {
                ListItem item = listItems[0];
                FieldLookupValue lookupValue = new FieldLookupValue();
                lookupValue.LookupId = Convert.ToInt32(item["ID"]);
                return lookupValue;
            }
            else
            {
                ListItemCreationInformation li = new ListItemCreationInformation();
                ListItem oListItem = lookupList.AddItem(li);
                
                //Set each field value
                oListItem[lookupFieldName] = lookupName;
                //Persist changes
                oListItem.Update();
                clientContext.ExecuteQuery();

                FieldLookupValue lookupValue = new FieldLookupValue();
                lookupValue.LookupId = oListItem.Id;
                return lookupValue;
            }
        }

        public static TaxonomyFieldValue GetTaxonomyFieldValue(string termName, Microsoft.SharePoint.Client.Field mmField, ClientContext clientContext)
        {
            //Cast field to TaxonomyField to get its TermSetId
            TaxonomyField taxField = clientContext.CastTo<TaxonomyField>(mmField);
            //Get term ID from name and term set ID
            string termId = GetTermIdForTerm(termName, taxField.TermSetId, clientContext);
            if (!string.IsNullOrEmpty(termId))
            {
                //Set TaxonomyFieldValue
                TaxonomyFieldValue termValue = new TaxonomyFieldValue();
                termValue.Label = termName;
                termValue.TermGuid = termId;
                termValue.WssId = -1;
                return termValue;
            }
            return null;
        }

        public static string GetTermIdForTerm(string term, Guid termSetId, ClientContext clientContext)
        {
            string termId = string.Empty;

            //Get term set from ID
            TaxonomySession tSession = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore ts = tSession.GetDefaultSiteCollectionTermStore();
            TermSet tset = ts.GetTermSet(termSetId);

            LabelMatchInformation lmi = new LabelMatchInformation(clientContext);

            lmi.Lcid = 1033;
            lmi.TrimUnavailable = true;
            lmi.TermLabel = term;

            //Search for matching terms in the term set based on label
            TermCollection termMatches = tset.GetTerms(lmi);
            clientContext.Load(tSession);
            clientContext.Load(ts);
            clientContext.Load(tset);
            clientContext.Load(termMatches);

            clientContext.ExecuteQuery();

            //Set term ID to first match
            if (termMatches != null && termMatches.Count() > 0)
                termId = termMatches.First().Id.ToString();

            return termId;
        }
    }
}
