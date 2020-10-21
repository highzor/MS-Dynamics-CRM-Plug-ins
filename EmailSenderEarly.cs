using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Crm.Sdk.Samples.EarlyBoundTypes;

namespace Microsoft.Crm.Sdk.Samples
{
    public class EmailSenderEarly : IPlugin
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
                    Contact contact = GetContact(service, context);
                    SendEmailRequest sendEmailRequest = CreateEmail(contact, service, context);
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
        private Contact GetContact(IOrganizationService service, IPluginExecutionContext context)
        {
            Contact contact = (Contact)service.Retrieve(Contact.EntityLogicalName, context.PrimaryEntityId, new ColumnSet("fullname", "new_regionfield", "new_cityfield", "emailaddress1"));
            return contact;
        }
        private SendEmailRequest CreateEmail(Contact contact, IOrganizationService service, IPluginExecutionContext context)
        {
            ActivityParty fromParty = new ActivityParty
            {
                PartyId = new EntityReference(SystemUser.EntityLogicalName, context.UserId)
            };
            ActivityParty toParty = new ActivityParty
            {
                PartyId = new EntityReference(Contact.EntityLogicalName, contact.Id)
            };
            Email email = new Email
            {
                To = new ActivityParty[] { toParty },
                From = new ActivityParty[] { fromParty },
                Subject = "SDK Sample e-mail early-bound style",
                Description = $"Hello: {contact.FullName}, {contact.Region.Name}, {contact.City.Name}, {contact.Email}",
                DirectionCode = true
            };
            Guid emailId = service.Create(email);
            SendEmailRequest sendEmailreq = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = "",
                IssueSend = true
            };
            return sendEmailreq;
        }
    }
}
