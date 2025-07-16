using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.Collections;
using System.Xml.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Data;
using DataverseModel;

namespace ContactPreferencePlugin
{
    public class ContactPreferencePlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService) serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext) serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory) serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target"))
            {
                int stageNumber = 1;
                Entity entity = (Entity)context.InputParameters["Target"];
                tracingService.Trace($"Stage {stageNumber++}: Retrieved target entity");
                
                Entity preImage = null;
                Contact contact = null;
                if (context.MessageName == "Update")
                {
                    if (context.PreEntityImages.Contains("PreImage"))
                    {
                        preImage = context.PreEntityImages["PreImage"];
                    }

                    contact = preImage?.ToEntity<Contact>();
                }
                else if (context.MessageName == "Create")
                {
                    contact = entity.ToEntity<Contact>();
                }
                if (contact != null)
                {
                    tracingService.Trace($"Stage {stageNumber++}: Retrieved contact");
                    var query = new QueryExpression(Contact.EntityLogicalName)
                    {
                        ColumnSet = new ColumnSet(Contact.Fields.DoNotEmail),
                        Criteria =
                        {
                            Conditions =
                            {
                                new ConditionExpression(Contact.Fields.ParentCustomerId, ConditionOperator.Equal, contact.ParentCustomerId.Id),
                                new ConditionExpression(Contact.Fields.DoNotEmail, ConditionOperator.Equal, false)
                            }
                        }
                    };
                    var contactCount= service.RetrieveMultiple(query).Entities.Count;
                    tracingService.Trace($"{contactCount} contacts without 'Do Not Email' preference found.");

                    var updateAccount = new Account { Id = contact.ParentCustomerId.Id };
                    if (contactCount == 1)
                    {
                        tracingService.Trace($"Stage {stageNumber++}: Only 1 contact reachable via email for related account. Account risk level is moderate.");
                        updateAccount.CR950_RiskLevel = Account_CR950_RiskLevel.Moderate;
                    }
                    else if (contactCount < 1)
                    {
                        tracingService.Trace($"Stage {stageNumber++}: No contacts reachable via email for related account.  Account risk level is high.");
                        updateAccount.CR950_RiskLevel = Account_CR950_RiskLevel.High;
                    }
                    else
                    {
                        tracingService.Trace($"Multiple contacts reachable via email for related account. Account risk level is normal.");
                        updateAccount.CR950_RiskLevel = Account_CR950_RiskLevel.Low;
                    }
                    try
                    {
                        service.Update(updateAccount);
                        tracingService.Trace($"Stage {stageNumber++}: Account update succeeded.");
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace($"Stage {stageNumber++}: Error updating account - {ex.ToString()}");
                        throw;
                    }

                }
                else
                {
                    tracingService.Trace($"Stage {stageNumber++}: Could not get contact");
                }

            }
        }
    }
}
