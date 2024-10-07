using System;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using System.Text;
using OfficeOpenXml;
using System.Linq;

namespace CRMTool.Services
{
    internal class EntityMetadataService
    {
        private IOrganizationService _service;
        private string orgUrl;
        private string accessToken;

        public EntityMetadataService(IOrganizationService service)
        {
            _service = service;
            orgUrl = new Uri(((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).CrmConnectOrgUriActual.OriginalString).GetLeftPart(UriPartial.Authority).Replace(".api", "");
            accessToken = ((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).CurrentAccessToken;
        }

        public async Task<string> callCRMAPI(string endpoint)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    string requestUrl = $"{orgUrl}/api/data/v9.2/{endpoint}";

                    HttpResponseMessage response = await client.GetAsync(requestUrl);

                    if (response.IsSuccessStatusCode)
                    {
                        return await response.Content.ReadAsStringAsync(); // Return the response content (metadata)
                    }
                    else
                    {
                        Console.WriteLine($"Failed to fetch metadata. Status Code: {response.StatusCode}");
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching metadata: {ex.Message}");
                return null;
            }
        }



        public async Task FetchAndSaveFullEntityMetadataAsync(string entityid, string filePath)
        {
            string endpoint = $"EntityDefinitions({entityid})";

            try
            {
                var fullMetadataJson = await callCRMAPI(endpoint);

                if (!string.IsNullOrEmpty(fullMetadataJson))
                {
                    var parsedJson = JsonConvert.DeserializeObject<Dictionary<string, object>>(fullMetadataJson);

                    string primaryColumnLogicalName = GetValueOrDefault(parsedJson, "PrimaryNameAttribute");
                    string entityLogicalName = GetValueOrDefault(parsedJson, "LogicalName");
                    var primaryColumnMetadataJson = await FetchPrimaryColumnMetadataAsync(primaryColumnLogicalName, entityLogicalName);
                    var primaryColumnMetadata = JsonConvert.DeserializeObject<Dictionary<string, object>>(primaryColumnMetadataJson);



                    // Helper function to safely extract values from the dictionary
                    string GetValueOrDefault(Dictionary<string, object> dict, string key, string defaultValue = "")
                    {
                        if (dict.ContainsKey(key) && dict[key] != null)
                        {
                            var value = dict[key];

                            if (value is JObject obj)
                            {
                                if (obj.ContainsKey("Value"))
                                {
                                    return obj["Value"].ToString();
                                }

                                if (obj.ContainsKey("LocalizedLabels") && obj["LocalizedLabels"] is JArray labelsArray)
                                {
                                    var labelObj = labelsArray.FirstOrDefault();
                                    if (labelObj != null && labelObj["Label"] != null)
                                    {
                                        return labelObj["Label"].ToString();
                                    }
                                }
                            }
                            else
                            {
                                return value.ToString();
                            }
                        }
                        return defaultValue;
                    }

                    // Prepare the data to append
                    var data = new List<object[]>
                    {
                        new object[]
                        {
                            GetValueOrDefault(parsedJson, "DisplayName"), // Display name
                            GetValueOrDefault(parsedJson, "PluralName"), // Plural name
                            GetValueOrDefault(parsedJson, "Description"), // Description
                            GetValueOrDefault(parsedJson, "HasNotes"), // Enable attachments
                            GetValueOrDefault(parsedJson, "SchemaName"), // Schema name
                            GetValueOrDefault(parsedJson, "LogicalName"), // Logical name
                            GetValueOrDefault(parsedJson, "TableType"), // Type
                            GetValueOrDefault(parsedJson, "OwnershipType"), // Record ownership
                            GetValueOrDefault(parsedJson, "ChooseTableImage"), // Choose table image
                            GetValueOrDefault(parsedJson, "EntityColor"), // Color
                            GetValueOrDefault(parsedJson, "IsDuplicateDetectionEnabled"), // Apply duplicate detection rules
                            GetValueOrDefault(parsedJson, "ChangeTrackingEnabled"), // Track changes
                            GetValueOrDefault(parsedJson, "EntityHelpUrlEnabled"), // Provide custom help
                            GetValueOrDefault(parsedJson, "EntityHelpUrl"), // Help URL
                            GetValueOrDefault(parsedJson, "AuditChangesToItsData"), // Audit changes to its data
                            GetValueOrDefault(parsedJson, "IsQuickCreateEnabled"), // Leverage quick-create form if available
                            GetValueOrDefault(parsedJson, "IsRetentionEnabled"), // Enable long term retention
                            GetValueOrDefault(parsedJson, "EnableRecycleBin"), // Enable recycle bin
                            GetValueOrDefault(parsedJson, "HasActivities"), // Creating a new activity
                            GetValueOrDefault(parsedJson, "IsMailMergeEnabled"), // Doing a mail merge
                            GetValueOrDefault(parsedJson, "SettingUpSharePointDocumentManagement"), // Setting up SharePoint document management
                            GetValueOrDefault(parsedJson, "IsConnectionsEnabled"), // Can have connections
                            GetValueOrDefault(parsedJson, "HasEmailAddresses"), // Can have a contact email
                            GetValueOrDefault(parsedJson, "HaveAnAccessTeam"), // Have an access team
                            GetValueOrDefault(parsedJson, "HasFeedback"), // Can be linked to feedback
                            GetValueOrDefault(parsedJson, "SyncToExternalSearchIndex"), // Appear in search results
                            GetValueOrDefault(parsedJson, "CanBeTakenOffline"), // Can be taken offline
                            GetValueOrDefault(parsedJson, "IsValidForQueue"), // Can be added to a queue
                            GetValueOrDefault(parsedJson, "AutoRouteToOwnerQueue"), // When rows are created or assigned, move them to the owner's default queue
                             // Primary column metadata
                            GetValueOrDefault(primaryColumnMetadata, "DisplayName"), // Primary column display name
                            GetValueOrDefault(primaryColumnMetadata, "Description"), // Primary column description
                            GetValueOrDefault(primaryColumnMetadata, "SchemaName"), // Primary column schema name
                            GetValueOrDefault(primaryColumnMetadata, "LogicalName"), // Primary column logical name
                            GetValueOrDefault(primaryColumnMetadata, "RequiredLevel"), // Primary column requirement
                            GetValueOrDefault(primaryColumnMetadata, "MaxLength"), // Primary column maximum character count

                           


                        }
                    };

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    FileInfo excelFile = new FileInfo(filePath);
                    string newLogicalName = data[0][5].ToString(); // Logical name is the 6th element in the new data

                    // Check if the file exists
                    if (excelFile.Exists)
                    {
                        using (var package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("Entity Metadata");

                            // Check for existing logical names to avoid duplicates
                            bool logicalNameExists = false;
                            int rowCount = worksheet.Dimension?.Rows ?? 0;

                            for (int row = 2; row <= rowCount; row++) // Start from 2 to skip header row
                            {
                                string existingLogicalName = worksheet.Cells[row, 6].Text; // Logical Name is in column 6
                                if (existingLogicalName.Equals(newLogicalName, StringComparison.OrdinalIgnoreCase))
                                {
                                    logicalNameExists = true;
                                    break;
                                }
                            }

                            // If logical name does not exist, append the data
                            if (!logicalNameExists)
                            {
                                var startRow = rowCount + 1; // Next available row
                                worksheet.Cells[startRow, 1].LoadFromArrays(data);
                                package.Save();
                                Console.WriteLine("Entity metadata appended successfully to the Excel file.");
                            }
                            else
                            {
                                Console.WriteLine($"Logical name '{newLogicalName}' already exists. Skipping append operation.");
                            }
                        }
                    }
                    else
                    {
                        using (var package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Entity Metadata");

                            // Add headers
                            worksheet.Cells["A1"].LoadFromArrays(new List<object[]>()
                    {
                        new object[]
                        {
                            "Display name", "Plural name", "Description", "Enable attachments", "Schema name",
                            "Logical name", "Type", "Record ownership", "Choose table image", "Color",
                            "Apply duplicate detection rules", "Track changes", "Provide custom help",
                            "Help URL", "Audit changes to its data", "Leverage quick-create form if available",
                            "Enable long term retention", "Enable recycle bin", "Creating a new activity",
                            "Doing a mail merge", "Setting up SharePoint document management", "Can have connections",
                            "Can have a contact email", "Have an access team", "Can be linked to feedback",
                            "Appear in search results", "Can be taken offline", "Can be added to a queue",
                            "When rows are created or assigned, move them to the owner's default queue",
                            "Primary column display name", "Primary column description",
                            "Primary column schema name", "Primary column logical name",
                            "Primary column requirement", "Primary column maximum character count"
                        }
                    });

                            // Append the data
                            worksheet.Cells[2, 1].LoadFromArrays(data);
                            package.Save();
                            Console.WriteLine("Entity metadata saved successfully to the new Excel file.");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No metadata returned from the API.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching or saving entity metadata: {ex.Message}");
            }
        }


        public async Task<string> FetchPrimaryColumnMetadataAsync(string primaryColumnLogicalName, string entityLogicalName)
        {
            // Define the endpoint to fetch primary column metadata for the specified entity
            string endpoint = $"EntityDefinitions(LogicalName='{entityLogicalName}')/Attributes(LogicalName='{primaryColumnLogicalName}')";

            try
            {
                // Call the CRM API and await the response
                var response = await callCRMAPI(endpoint);

                // Check if a valid response was returned
                if (!string.IsNullOrEmpty(response))
                {
                    // Return the response if successful
                    return response;
                }
                else
                {
                    // Log a message if no data was returned
                    Console.WriteLine("No data returned from the API.");
                    return null; // Return null to indicate no data
                }
            }
            catch (Exception ex)
            {
                // Catch any exceptions and log an error message
                Console.WriteLine($"Error fetching AttributeMetaData: {ex.Message}");
                return null; // Return null in case of an error
            }
        }


        public async Task FetchEntityMetadataAsync(string prefix)
        {
            string endpoint = $"entities?$select=entityid,logicalname,originallocalizedname&$filter=startswith(logicalname,'{prefix}')";
            string filePath = @"C:\Users\rajpo\Documents\entity_metadata.json";

            try
            {
                var entityDataJson = await callCRMAPI(endpoint);

                if (entityDataJson != null)
                {
                    var entityMetadataResponse = JsonConvert.DeserializeObject<EntityMetadataResponse>(entityDataJson);

                    var formattedEntities = new List<Dictionary<string, string>>();

                    foreach (var entity in entityMetadataResponse.Value)
                    {
                        var entityDict = new Dictionary<string, string>
                        {
                            { "Entity Id", entity.EntityId },
                            { "Logical Name", entity.LogicalName },
                            { "Display Name", entity.OriginalLocalizedName }
                        };
                        formattedEntities.Add(entityDict);
                    }

                    string formattedJsonData = JsonConvert.SerializeObject(formattedEntities, Formatting.Indented);
                    File.WriteAllText(filePath, formattedJsonData);

                    Console.WriteLine("Data saved successfully to the file.");
                }
                else
                {
                    Console.WriteLine("No data returned from the API.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching or saving data: {ex.Message}");
            }
        }
    }

    public class Metadata
    {
        public string EntityId { get; set; }
        public string LogicalName { get; set; }
        public string OriginalLocalizedName { get; set; }
    }

    public class EntityMetadataResponse
    {
        public List<Metadata> Value { get; set; }
    }

}


