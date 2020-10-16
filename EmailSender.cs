using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.ServiceModel;
namespace Microsoft.Crm.Sdk.Samples
{
    [CrmPluginRegistration(
        message: MessageNameEnum.Create,
        entityLogicalName: "contact",
        stage: StageEnum.PostOperation,
        executionMode: ExecutionModeEnum.Synchronous,
        filteringAttributes: null,
        stepName: "EmailSenderCreate",
        executionOrder: 1,
        isolationModel: IsolationModeEnum.Sandbox,
        Id = "6CE2094C-EBA9-4E85-9DA7-4F86A83F9603")]

    public class EmailSender : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                try
                {
                    Entity _Contact = CreateContact(service, context);
                    SendEmailRequest sendEmailRequest = CreateEmail(_Contact, context, service);
                    SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException($"FaultException exception: {ex.Message}");
                }
                catch (Exception ex)
                {
                    tracingService.Trace($"FollowUpPlugin: {ex.Message}");
                    throw new InvalidPluginExecutionException($"An exception happend: {ex.Message}");
                }
            }
        }
        public Entity CreateContact(IOrganizationService service, IPluginExecutionContext context)
        {
            Entity _Contact = service.Retrieve("contact", context.PrimaryEntityId,
                        new ColumnSet("fullname", "new_regionfield", "new_cityfield", "emailaddress1"));
            return _Contact;
        }
        public SendEmailRequest CreateEmail(Entity _Contact, IPluginExecutionContext context, IOrganizationService service)
        {
            string Name = "";
            if (_Contact["fullname"] != null)
            {
                Name = _Contact["fullname"].ToString();
            }
            EntityReference Region = _Contact["new_regionfield"] as EntityReference;
            EntityReference City = _Contact["new_cityfield"] as EntityReference;
            string Email = "";
            if (_Contact["emailaddress1"] != null)
            {
                Email = _Contact["emailaddress1"].ToString();
            }
            Entity fromActivityParty = new Entity("activityparty");
            Entity toActivityParty = new Entity("activityparty");
            Guid userId = context.UserId;
            Guid contactId = (Guid)_Contact.Attributes["contactid"];
            fromActivityParty["partyid"] = new EntityReference("systemuser", userId);
            toActivityParty["partyid"] = new EntityReference("contact", contactId);
            Entity email = new Entity("email");
            email["from"] = new Entity[] { fromActivityParty };
            email["to"] = new Entity[] { toActivityParty };
            email["regardingobjectid"] = new EntityReference("contact", contactId);
            email["subject"] = "This is the subject";
            email["description"] = $"Hello: {Name}, {Region.Name}, {City.Name}, {Email}";
            email["directioncode"] = true;
            Guid emailId = service.Create(email);
            SendEmailRequest sendEmailRequest = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = "",
                IssueSend = true
            };
            return sendEmailRequest;
        }
    }
}
