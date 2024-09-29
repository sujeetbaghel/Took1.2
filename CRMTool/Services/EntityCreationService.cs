using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;



namespace CRMTool.Services
{
    public class EntityCreationService
    {
        private readonly IOrganizationService _service;
        private readonly MyPluginControl _myPluginControl; // Reference to the UI control


        public EntityCreationService(IOrganizationService service, MyPluginControl myPluginControl)
        {
            _service = service;
            _myPluginControl = myPluginControl; // Store the reference

        }


        public async Task CreateCustomEntityAsync(Dictionary<string, string> entityProperties)
        {
            try
            {
                string entityName = entityProperties["schema name"];
                if (string.IsNullOrWhiteSpace(entityName))
                {
                    throw new ArgumentException("Schema Name cannot be empty.");
                }

                EntityMetadata newEntity = new EntityMetadata
                {
                    SchemaName = entityName,
                    DisplayName = new Label(entityProperties["display name"], 1033),
                    LogicalName = entityName,
                    Description = new Label(entityProperties["description"], 1033),
                    IsVisibleInMobile = new BooleanManagedProperty(true),
                    IsVisibleInMobileClient = new BooleanManagedProperty(true),
                    OwnershipType = OwnershipTypes.UserOwned,
                    DisplayCollectionName = new Label(entityProperties["plural name"], 1033),
                    //IsActivity = entityProperties["type"] == "Activity"

                };

                var nameAttribute = new StringAttributeMetadata
                {
                    SchemaName = entityProperties["primary column schema name"],
                    LogicalName = entityProperties["primary column logical name"],
                    DisplayName = new Label(entityProperties["primary column display name"], 1033),
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MaxLength = int.TryParse(entityProperties["primary column maximum character count"], out int maxLength) ? maxLength : 100,
                    Description = new Label(entityProperties["primary column description"], 1033)
                };

                CreateEntityRequest createEntityRequest = new CreateEntityRequest
                {
                    Entity = newEntity,
                    PrimaryAttribute = nameAttribute,
                };

                var response = (CreateEntityResponse)_service.Execute(createEntityRequest);

                _myPluginControl.UpdateLog($"Entity '{entityName}' created successfully with ID: {response.EntityId}");
            }
            catch (Exception ex)
            {
                _myPluginControl.UpdateLog($"Error creating entity '{entityProperties["schema name"]}' : {ex.Message}");
            }

        }
    }
}
