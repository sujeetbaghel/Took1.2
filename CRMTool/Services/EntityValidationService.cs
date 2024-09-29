using System;
using System.Collections.Generic;
using System.Linq;

namespace CRMTool.Services
{
    public class EntityValidationService
    {
        // Dictionary to hold field names and their expected data types, case-insensitive
        private readonly Dictionary<string, string> fieldTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "display name", "String" },
            { "plural name", "String" },
            { "description", "String" },
            { "enable attachments", "Boolean" },
            { "schema name", "String" },
            { "logical name", "String" },
            { "type", "String" },
            { "record ownership", "Integer" },
            { "choose table image", "String" },
            { "color", "String" },
            { "apply duplicate detection rules", "Boolean" },
            { "track changes", "Boolean" },
            { "provide custom help", "Boolean" },
            { "help url", "String" },
            { "audit changes to its data", "Boolean" },
            { "leverage quick-create form if available", "Boolean" },
            { "enable long term retention", "Boolean" },
            { "enable recycle bin", "Boolean" },
            { "creating a new activity", "Boolean" },
            { "doing a mail merge", "Boolean" },
            { "setting up sharepoint document management", "Boolean" },
            { "can have connections", "Boolean" },
            { "can have a contact email", "Boolean" },
            { "have an access team", "Boolean" },
            { "can be linked to feedback", "Boolean" },
            { "appear in search results", "Boolean" },
            { "can be taken offline", "Boolean" },
            { "can be added to a queue", "Boolean" },
            { "when rows are created or assigned, move them to the owner's default queue", "Boolean" },
            { "primary column display name", "String" },
            { "primary column description", "String" },
            { "primary column schema name", "String" },
            { "primary column logical name", "String" },
            { "primary column requirement", "Integer" },
            { "primary column maximum character count", "Integer" }
        };

        public bool ValidateEntityProperties(Dictionary<string, string> entityProperties, out List<string> errorMessages)
        {
            errorMessages = new List<string>();


            // List of required fields
            List<string> requiredFields = new List<string>(fieldTypes.Keys);

            foreach (var field in requiredFields)
            {
                // Check if the field exists and is not empty
                if (!entityProperties.ContainsKey(field) || string.IsNullOrWhiteSpace(entityProperties[field]))
                {
                    var schemaName = !string.IsNullOrWhiteSpace(entityProperties["schema name"]) ? entityProperties["schema name"] : "unknown schema";
                    errorMessages.Add($"{field} is missing or empty for entity: {schemaName}");
                    continue; // Proceed to check the next field
                }

                // Validate boolean fields
                if (fieldTypes[field] == "Boolean" && !new[] { "true", "false", "1", "0", "yes", "no" }.Contains(entityProperties[field].ToLower()))
                {
                    var schemaName = !string.IsNullOrWhiteSpace(entityProperties["schema name"]) ? entityProperties["schema name"] : "unknown schema";
                    errorMessages.Add($"{field} should be either 'true' or 'false' for entity: {schemaName}");
                    continue; // Proceed to check the next field
                }

                // Validate integer fields
                if (fieldTypes[field] == "Integer" && !int.TryParse(entityProperties[field], out _))
                {
                    var schemaName = !string.IsNullOrWhiteSpace(entityProperties["schema name"]) ? entityProperties["schema name"] : "unknown schema";
                    errorMessages.Add($"{field} should be a valid integer for entity: {schemaName}");
                    continue; // Proceed to check the next field
                }
            }
            return errorMessages.Count == 0;
        }

    }
}
