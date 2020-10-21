using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using System.Configuration;
using Microsoft.Crm.Sdk.Samples.EarlyBoundTypes;

namespace Microsoft.Crm.Sdk.Samples
{
    public class EmailSenderEarly : IPlugin
    {
        private Guid _emailId;
        private Guid _contactId;
        private Guid _userId;
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
                       Contact contact = CreateContact(service, context);
                        SendEmailRequest sendEmailRequest = CreateEmail(contact, service);
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
        private Contact CreateContact(IOrganizationService service, IPluginExecutionContext context)
        {            
            Contact contact = (Contact)service.Retrieve(Contact.EntityLogicalName, context.PrimaryEntityId, new ColumnSet("fullname", "new_regionfield", "new_cityfield", "emailaddress1"));
            return contact;
        }
        private SendEmailRequest CreateEmail(Contact contact, IOrganizationService service)
        {
                WhoAmIRequest systemUserRequest = new WhoAmIRequest();
                WhoAmIResponse systemUserResponse = (WhoAmIResponse)service.Execute(systemUserRequest);
                _userId = systemUserResponse.UserId;
                _contactId = (Guid)contact.ContactId;
                ActivityParty fromParty = new ActivityParty
                {
                    PartyId = new EntityReference(SystemUser.EntityLogicalName, _userId)
                };
                ActivityParty toParty = new ActivityParty
                {
                    PartyId = new EntityReference(Contact.EntityLogicalName, _contactId)
                };
                Email email = new Email
                {
                    To = new ActivityParty[] { toParty },
                    From = new ActivityParty[] { fromParty },
                    Subject = "SDK Sample e-mail early-bound style",
                    Description = $"Hello: {contact.FullName}, {contact.new_regionfield.Name}, {contact.new_cityfield.Name}, {contact.EMailAddress1}",
                    DirectionCode = true
                };
                _emailId = service.Create(email);
                SendEmailRequest sendEmailreq = new SendEmailRequest
                {
                    EmailId = _emailId,
                    TrackingToken = "",
                    IssueSend = true
                };
                return sendEmailreq;
        }
    }
}
