using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace ColorfullFlagIconDynamics365Engine
{
    public class Register : Plugin
    {
        public const string ASSEMBLY_NAME = "ColorfullFlagIconDynamics365Engine";
        public const string GENARAL_LOGIC_PLUGIN_TYPE = "ColorfullFlagIconDynamics365Engine.FlagIconLogicExecutorCreateUpdate";

        public Register() : base(typeof(Register))
        {
            base.RegisteredEvents.Add(new Tuple<int, string, string, Action<LocalPluginContext>>(20, "Create", "clfi_configurationentity", new Action<LocalPluginContext>(create)));

            base.RegisteredEvents.Add(new Tuple<int, string, string, Action<LocalPluginContext>>(20, "Update", "clfi_configurationentity", new Action<LocalPluginContext>(update)));

            base.RegisteredEvents.Add(new Tuple<int, string, string, Action<LocalPluginContext>>(10, "Delete", "clfi_configurationentity", new Action<LocalPluginContext>(delete)));

            base.RegisteredEvents.Add(new Tuple<int, string, string, Action<LocalPluginContext>>(40, "Update", "clfi_iconstatusconfiguration", new Action<LocalPluginContext>(bulkUpdateStatusIcon)));

        }
        protected void create(LocalPluginContext localContext)
        {
            IPluginExecutionContext context = localContext.PluginExecutionContext;
            
            IOrganizationService service = localContext.OrganizationService;

            try
            {
                var currentEntity = context.InputParameters["Target"] as Entity; // Configuration Entity


                
                  if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        string entityLogicalName = currentEntity.GetAttributeValue<string>("clfi_entitylogicalname").ToLower();
                        string iconLogicaFieldName = currentEntity.GetAttributeValue<string>("clfi_iconlogicalfieldname").ToLower();
                        string statusLogicaFieldName = currentEntity.GetAttributeValue<string>("clfi_statuslogicalfieldname").ToLower();
                        
                        

                        // verify if already exists 
                        if(configurationExists(entityLogicalName, service))
                        {
                            throw new InvalidPluginExecutionException("ERROR : The is already a configuration involving the same entity logical name");
                        }
                        else
                        {
                            
                            Guid stepIdCreate = SdkMessageStep(ASSEMBLY_NAME, GENARAL_LOGIC_PLUGIN_TYPE, service, "Create", entityLogicalName, 0, 40);
                            Guid stepIdUpdate = SdkMessageStep(ASSEMBLY_NAME, GENARAL_LOGIC_PLUGIN_TYPE, service, "Update", entityLogicalName, 0, 40, statusLogicaFieldName);
                        } 
                    }
               
            }
            catch (InvalidPluginExecutionException ex)
            {
                throw ex;
            }
        }
        protected void update(LocalPluginContext localContext)
        {
            IPluginExecutionContext context = localContext.PluginExecutionContext;

            IOrganizationService service = localContext.OrganizationService;

            try
            {
                var currentEntity = context.InputParameters["Target"] as Entity; // Configuration Entity


                
                    var currentConfig = service.Retrieve(currentEntity.LogicalName, currentEntity.Id, new ColumnSet("clfi_entitylogicalname", "clfi_iconlogicalfieldname", "clfi_statuslogicalfieldname"));
                    var entityLogicalName = currentConfig.GetAttributeValue<string>("clfi_entitylogicalname").ToLower();
                    var iconLogicaFieldName = currentConfig.GetAttributeValue<string>("clfi_iconlogicalfieldname").ToLower();
                    var statusLogicaFieldName = currentConfig.GetAttributeValue<string>("clfi_statuslogicalfieldname").ToLower();
                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        //Main function inside which other function has been called to add plugin step
                        Guid stepId = SdkMessageStep(ASSEMBLY_NAME, GENARAL_LOGIC_PLUGIN_TYPE, service, "Update", entityLogicalName, 0, 40, statusLogicaFieldName, "Update");
                    }
                
               
            }
            catch (InvalidPluginExecutionException ex)
            {
                throw ex;
            }
        }
        protected void delete(LocalPluginContext localContext)
        {
            IPluginExecutionContext context = localContext.PluginExecutionContext;

            IOrganizationService service = localContext.OrganizationService;

            try
            {
                var currentEntity = (EntityReference)context.InputParameters["Target"] ; // Configuration Entity
                var currentConfig = service.Retrieve(currentEntity.LogicalName, currentEntity.Id, new ColumnSet("clfi_entitylogicalname", "clfi_iconlogicalfieldname", "clfi_statuslogicalfieldname"));


                string entityLogicalName = currentConfig.GetAttributeValue<string>("clfi_entitylogicalname").ToLower();
                    deleteConfiguration(ASSEMBLY_NAME, GENARAL_LOGIC_PLUGIN_TYPE, service, "Create", entityLogicalName);
                    deleteConfiguration(ASSEMBLY_NAME, GENARAL_LOGIC_PLUGIN_TYPE, service, "Update", entityLogicalName);
            }
            catch (InvalidPluginExecutionException ex)
            {
                throw ex;
            }
        }
        private static void deleteConfiguration(string assemblyName, string pluginTypeName, IOrganizationService service, string messageName, string entityLogicalName)
        {
            Guid pluginTypeId = GetPluginTypeId(assemblyName, pluginTypeName, service);
            Guid messageId = GetSdkMessageId(messageName, service);
            var stepToDelete = GetSdkMessageStepId(messageId, pluginTypeId, entityLogicalName, service);
            service.Delete(stepToDelete.LogicalName, stepToDelete.Id);

        }
        private static bool configurationExists(string entityLogicalName, IOrganizationService service)
        {
            QueryExpression queryConfig = new QueryExpression("clfi_configurationentity");
            queryConfig.Criteria = new FilterExpression
            {
                Conditions =
                            {
                                new ConditionExpression
                                {
                                  AttributeName = "clfi_entitylogicalname",
                                  Operator = ConditionOperator.Equal,
                                  Values = { entityLogicalName }
                                 },
                             }
            };
            EntityCollection configurations = service.RetrieveMultiple(queryConfig);


            // verify if already exists 
            if (configurations.Entities.Count > 0)
                return true;

            return false;
        }
        private static void bulkUpdateStatusIcon(LocalPluginContext localContext)
        {
            IPluginExecutionContext context = localContext.PluginExecutionContext;

            IOrganizationService service = localContext.OrganizationService;

            try
            {   Entity imageEntity = null;
                if (context.PreEntityImages.Contains("PreImage"))
                {
                    imageEntity = context.PreEntityImages["PreImage"];
                }
                var tempCurrent = context.InputParameters["Target"] as Entity;
                var currentEntity = service.Retrieve(tempCurrent.LogicalName, tempCurrent.Id, new ColumnSet("clfi_statusvalue", "clfi_icon")); // Configuration Icon Status Entity
                var currentConfig = service.Retrieve("clfi_configurationentity", imageEntity.GetAttributeValue<EntityReference>("clfi_targetentity").Id,new ColumnSet(true));
                var entityLogicalName = currentConfig.GetAttributeValue<string>("clfi_entitylogicalname").ToLower();
                var iconLogicaFieldName = currentConfig.GetAttributeValue<string>("clfi_iconlogicalfieldname").ToLower();
                var statusLogicaFieldName = currentConfig.GetAttributeValue<string>("clfi_statuslogicalfieldname").ToLower();
                
                if (currentEntity.Contains("clfi_icon") && currentEntity.Contains("clfi_statusvalue"))
                {
                    var statusValue = currentEntity.GetAttributeValue<int>("clfi_statusvalue");
                    var newIcon = currentEntity["clfi_icon"] as byte[];

                    var queryEntityLogicalName = new QueryExpression(entityLogicalName);
                    queryEntityLogicalName.ColumnSet = new ColumnSet(statusLogicaFieldName, iconLogicaFieldName);
                    queryEntityLogicalName.Criteria = new FilterExpression
                    {
                        Conditions =
                              {
                                    new ConditionExpression
                                    {
                                      AttributeName = statusLogicaFieldName, //Configuration Entity  : lookup
                                      Operator = ConditionOperator.Equal,
                                      Values = { statusValue }
                                     },
                               }
                    };
                    List<Entity> logicalNameEntities = new List<Entity>();
                    getRecordByQuery(service,queryEntityLogicalName, ref logicalNameEntities);


                    var multipleRequest = new ExecuteMultipleRequest()
                    {
                        // Assign settings that define execution behavior: continue on error, return responses. 
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = true,
                            ReturnResponses = true
                        },
                        // Create an empty organization request collection.
                        Requests = new OrganizationRequestCollection()
                    };

                    // Add a UpdateRequest for each entity to the request collection.
                    ExecuteMultipleResponse multipleResponse = null;
                    foreach (var entity in logicalNameEntities)
                    {
                        UpdateRequest updateRequest = new UpdateRequest();
                        entity.Attributes.Remove(statusLogicaFieldName);
                        entity[iconLogicaFieldName] = newIcon;
                        updateRequest.Target = entity;
                        multipleRequest.Requests.Add(updateRequest);

                        if (multipleRequest
                            .Requests.Count == 1000)
                        {

                            multipleResponse = (ExecuteMultipleResponse)service.Execute(multipleRequest);
                            multipleRequest.Requests = new OrganizationRequestCollection();
                           
                        }
                    }

                    // Execute all the requests in the request collection using a single web method call.
                    if (multipleRequest.Requests.Count > 0)
                    {
                        multipleResponse = (ExecuteMultipleResponse)service.Execute(multipleRequest);
                        
                    }

                }


            }
            catch (InvalidPluginExecutionException ex)
            {
                throw ex;
            }
        }
        public static void getRecordByQuery(IOrganizationService service, QueryExpression query, ref List<Entity> result)
        {
            
            
            int queryCount = 500;
            int pageNumber = 1;

            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = queryCount;
            query.PageInfo.PageNumber = pageNumber;

            query.PageInfo.PagingCookie = null;

            while (true)
            {
                var queryRef = service.RetrieveMultiple(query);
                if (queryRef.Entities != null)
                {
                    result.AddRange(queryRef.Entities);
                }

                if (queryRef.MoreRecords)
                {
                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = queryRef.PagingCookie;
                }
                else
                {
                    break;
                }
            }

        }
        public Guid SdkMessageStep(string assemblyName, string pluginTypeName, IOrganizationService service, string messageName, string entityName, int mode, int stage,string filteringattributes =null, string eventName="Create")
        {
            Guid pluginTypeId = GetPluginTypeId(assemblyName, pluginTypeName, service);
            Guid messageId = GetSdkMessageId(messageName, service);
            Guid messageFitlerId = Guid.Empty;
            if (entityName != "" && entityName != string.Empty)
            {
                messageFitlerId = GetSdkMessageFilterId(entityName, messageName, service);
            }
            else
                entityName = "any entity";

            if ((pluginTypeId != Guid.Empty && pluginTypeId != null) && (messageId != null && messageId != Guid.Empty))
                {
                Guid stepId = Guid.Empty;
                if (eventName == "Update")
                {
                    

                    var tempStep = GetSdkMessageStepId(messageId,  pluginTypeId, entityName, service);
                     var stepToUpdate = new Entity(tempStep.LogicalName, tempStep.Id);
                    if (!string.IsNullOrEmpty(filteringattributes))
                        stepToUpdate["filteringattributes"] = filteringattributes;
                    service.Update(stepToUpdate);
                    return tempStep.Id;
                
                }
                else
                {
                    

                    Entity step = new Entity("sdkmessageprocessingstep");
                    step["name"] = pluginTypeName + ": " + messageName + " of " + entityName;
                    step["configuration"] = "";

                    if (!string.IsNullOrEmpty(filteringattributes))
                        step["filteringattributes"] = filteringattributes;
                    step["invocationsource"] = new OptionSetValue(0);
                    step["sdkmessageid"] = new EntityReference("sdkmessage", messageId);

                    step["supporteddeployment"] = new OptionSetValue(0);
                    step["plugintypeid"] = new EntityReference("plugintype", pluginTypeId);

                    step["mode"] = new OptionSetValue(mode); //0=sync,1=async
                    step["rank"] = 1;
                    step["stage"] = new OptionSetValue(stage); //10-preValidation, 20-preOperation, 40-PostOperation
                    if (messageFitlerId != null && messageFitlerId != Guid.Empty)
                    {
                        step["sdkmessagefilterid"] = new EntityReference("sdkmessagefilter", messageFitlerId);
                    }

                    stepId = service.Create(step);
                    return stepId;
                }
                

            }
            return Guid.Empty;
        }

        public static Guid GetPluginTypeId(string AssemblyName, string PluginTypeName, IOrganizationService service)
        {
            try
            {
                //GET ASSEMBLY QUERY
                QueryExpression pluginAssemblyQueryExpression = new QueryExpression("pluginassembly");
                pluginAssemblyQueryExpression.ColumnSet = new ColumnSet("pluginassemblyid");
                pluginAssemblyQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                      {
                            new ConditionExpression
                            {
                              AttributeName = "name",
                              Operator = ConditionOperator.Equal,
                              Values = { AssemblyName }
                             },
                        }
                };

                //RETRIEVE ASSEMBLY
                EntityCollection pluginAssemblies = service.RetrieveMultiple(pluginAssemblyQueryExpression);

                //IF ASSEMBLY IS FOUND
                if (pluginAssemblies.Entities.Count != 0)
                {
                    //ASSIGN ASSEMBLY ID TO VARIABLE
                    Guid assemblyId = pluginAssemblies.Entities.First().Id;

                    //GET PLUGIN TYPES WITHIN ASSEMBLY
                    QueryExpression pluginTypeQueryExpression = new QueryExpression("plugintype");
                    pluginTypeQueryExpression.ColumnSet = new ColumnSet("plugintypeid");
                    pluginTypeQueryExpression.Criteria = new FilterExpression
                    {
                        Conditions =
                                    {
                                        new ConditionExpression
                                        {
                                            AttributeName = "pluginassemblyid",
                                            Operator = ConditionOperator.Equal,
                                            Values = {assemblyId}
                                        },
                                        new ConditionExpression
                                        {
                                            AttributeName = "typename",
                                            Operator = ConditionOperator.Equal,
                                            Values = {PluginTypeName}
                                        },
                                    }
                    };

                    //RETRIEVE PLUGIN TYPES IN ASSEMBLY
                    EntityCollection pluginTypes = service.RetrieveMultiple(pluginTypeQueryExpression);

                    //RETURN PLUGIN TYPE ID
                    if (pluginTypes.Entities.Count != 0)
                    {

                        QueryExpression StepQueryExpression = new QueryExpression("sdkmessageprocessingstep");
                        StepQueryExpression.ColumnSet = new ColumnSet("name");
                        StepQueryExpression.Criteria = new FilterExpression
                        {
                            Conditions =
                                        {
                                            new ConditionExpression
                                            {
                                                AttributeName = "plugintypeid",
                                                Operator = ConditionOperator.Equal,
                                                Values = { pluginTypes.Entities.First().Id }
                                            }
                                        }
                        };

                        //RETRIEVE PLUGIN TYPES IN ASSEMBLY
                        //EntityCollection pluginSteps = service.RetrieveMultiple(StepQueryExpression);
                        //RETURN PLUGIN TYPE ID
                        if (pluginTypes.Entities.Count > 0)
                        {
                            return pluginTypes.Entities.First().Id;
                        }
                        else
                        {
                            return Guid.Empty;
                        }
                    }
                    else
                    {
                        return Guid.Empty;
                    }
                    throw new Exception(String.Format("Plugin Type {0} was not found in Assembly {1}", PluginTypeName, AssemblyName));
                }
                throw new Exception(String.Format("Assembly {0} not found", AssemblyName));
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
        public static Guid GetSdkMessageId(string SdkMessageName, IOrganizationService service)
        {
            try
            {
                //GET SDK MESSAGE QUERY
                QueryExpression sdkMessageQueryExpression = new QueryExpression("sdkmessage");
                sdkMessageQueryExpression.ColumnSet = new ColumnSet("sdkmessageid");
                sdkMessageQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                                {
                                    new ConditionExpression
                                    {
                                        AttributeName = "name",
                                        Operator = ConditionOperator.Equal,
                                        Values = {SdkMessageName}
                                    },
                                }
                };

                //RETRIEVE SDK MESSAGE
                EntityCollection sdkMessages = service.RetrieveMultiple(sdkMessageQueryExpression);
                if (sdkMessages.Entities.Count != 0)
                {
                    return sdkMessages.Entities.First().Id;
                }
                throw new Exception(String.Format("SDK MessageName {0} was not found.", SdkMessageName));
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public static Entity GetSdkMessageStepId(Guid sdkmessageid, Guid pluginTypeId, string entityLogicalName , IOrganizationService service)
        {
            try
            {
                //GET SDK STEP QUERY
                QueryExpression StepQueryExpression = new QueryExpression("sdkmessageprocessingstep");
                StepQueryExpression.ColumnSet = new ColumnSet(true);
                StepQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                                        {
                                            new ConditionExpression
                                            {
                                                AttributeName = "sdkmessageid",
                                                Operator = ConditionOperator.Equal,
                                                Values = { sdkmessageid }
                                            },

                                            new ConditionExpression
                                            {
                                                 AttributeName = "plugintypeid",
                                                Operator = ConditionOperator.Equal,
                                                Values = { pluginTypeId }
                                            },
                                            new ConditionExpression
                                            {
                                                AttributeName = "name",
                                                Operator = ConditionOperator.Like,
                                                Values = { "%" + entityLogicalName +"%" }
                                            },

                                        }
                };

                //RETRIEVE PLUGIN TYPES IN ASSEMBLY
                EntityCollection pluginSteps = service.RetrieveMultiple(StepQueryExpression);
                //RETURN PLUGIN TYPE ID
                if (pluginSteps.Entities.Count != 0)
                {
                    return pluginSteps.Entities.First();
                }
                else
                {
                    throw new Exception(String.Format("SDK MessageNameStep {0} was not found."));
                }
                
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public static Guid GetSdkMessageFilterId(string EntityLogicalName, string SdkMessageName, IOrganizationService service)
        {
            try
            {
                //GET SDK MESSAGE FILTER QUERY
                QueryExpression sdkMessageFilterQueryExpression = new QueryExpression("sdkmessagefilter");
                sdkMessageFilterQueryExpression.ColumnSet = new ColumnSet("sdkmessagefilterid");
                sdkMessageFilterQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                                {
                                    new ConditionExpression
                                    {
                                        AttributeName = "primaryobjecttypecode",
                                        Operator = ConditionOperator.Equal,
                                        Values = {EntityLogicalName}
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "sdkmessageid",
                                        Operator = ConditionOperator.Equal,
                                        Values = {GetSdkMessageId(SdkMessageName,service)}
                                    },
                                }
                };

                //RETRIEVE SDK MESSAGE FILTER
                EntityCollection sdkMessageFilters = service.RetrieveMultiple(sdkMessageFilterQueryExpression);

                if (sdkMessageFilters.Entities.Count != 0)
                {
                    return sdkMessageFilters.Entities.First().Id;
                }
                throw new Exception(String.Format("SDK Message Filter for {0} was not found.", EntityLogicalName));
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
    }
}
