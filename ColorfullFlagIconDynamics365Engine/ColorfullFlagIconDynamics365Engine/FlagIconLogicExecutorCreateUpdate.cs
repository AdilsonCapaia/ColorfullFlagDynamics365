using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace ColorfullFlagIconDynamics365Engine
{
    public class FlagIconLogicExecutorCreateUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

           
            try
            {
                var currentEntity = context.InputParameters["Target"] as Entity; // CDS entity

                //1- Query the Configuration entity to know wich field is the icon
                //Build a Configuration entity
                QueryExpression configurationEntityQueryExpression = new QueryExpression("clfi_configurationentity");
                configurationEntityQueryExpression.ColumnSet = new ColumnSet(true);
                configurationEntityQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                      {
                            new ConditionExpression
                            {
                              AttributeName = "clfi_entitylogicalname",
                              Operator = ConditionOperator.Equal,
                              Values = { currentEntity.LogicalName}
                             },
                        }
                };

                //RETRIEVE Configuration Entities
                EntityCollection ConfigurationEntities = service.RetrieveMultiple(configurationEntityQueryExpression);
                if (ConfigurationEntities.Entities.Count != 0)
                {
                    var concernedCFEntity = ConfigurationEntities.Entities.First();
                    //Name of the field that will contains the icon
                    var iconLogicaFieldName = concernedCFEntity.GetAttributeValue<string>("clfi_iconlogicalfieldname").ToLower();
                    //Name of the field to look for status value
                    var statusLogicaFieldName = concernedCFEntity.GetAttributeValue<string>("clfi_statuslogicalfieldname").ToLower();
                    //2- Check if field status exists and field icon exists on the CurrentEntity
                    if (currentEntity.Contains(statusLogicaFieldName))
                    {
                        //3- Query the Icon-Status configuration entity to get all the status-icon 
                        QueryExpression mapIconStatusEntityQueryExpression = new QueryExpression("clfi_iconstatusconfiguration");
                        mapIconStatusEntityQueryExpression.ColumnSet = new ColumnSet("clfi_statusvalue", "clfi_icon");
                        mapIconStatusEntityQueryExpression.Criteria = new FilterExpression
                        {
                            Conditions =
                              {
                                    new ConditionExpression
                                    {
                                      AttributeName = "clfi_targetentity", //Configuration Entity  : lookup
                                      Operator = ConditionOperator.Equal,
                                      Values = { concernedCFEntity.Id}
                                     },
                               }
                        };

                        //RETRIEVE Icon-Status records
                        EntityCollection IconStatusEntities = service.RetrieveMultiple(mapIconStatusEntityQueryExpression);
                        if (IconStatusEntities.Entities.Count != 0)
                        {
                            //4 if yes, proced to convert the icon value on the value type of entityimage 
                            foreach (var icoStatusRecord in IconStatusEntities.Entities)
                            {
                                if(currentEntity.GetAttributeValue<OptionSetValue>(statusLogicaFieldName)!= null && icoStatusRecord.GetAttributeValue<int>("clfi_statusvalue") == currentEntity.GetAttributeValue<OptionSetValue>(statusLogicaFieldName).Value)
                                {
                                    var entityToUpdate = new Entity(currentEntity.LogicalName, currentEntity.Id);
                                    //PB
                                    entityToUpdate[iconLogicaFieldName] = icoStatusRecord["clfi_icon"] as byte[] ;
                                    service.Update(entityToUpdate);
                                    
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("The 'Icon Logical field name' : " + iconLogicaFieldName + " or 'Status logical field name' : " + statusLogicaFieldName + " do not exis on the Entity logical name : " + currentEntity.LogicalName);
                    }
                }
                else
                {
                    throw new InvalidPluginExecutionException("There is not 'Configuration entity' for the entity :  " + currentEntity.LogicalName);
                }

            }
            catch(InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            
        }
    }
}
