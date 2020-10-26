using System.Linq;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Crm.Sdk.Samples.EarlyBoundTypes;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using System;

namespace Custom.WorkFlowAction1
{
    public class ChangeEmail : CodeActivity
    {
        [RequiredArgument]
        [Input("Email input")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> DescEmail { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);
            Email email = GetEmail(service, context);
            var partyInf = email.To.First();
            email.Description = $"<p><font size='5' color='red' face='Arial'>Ваша компания зарегистрирована {partyInf.PartyId.Name}. Будем рады продолжить сотрудничество.</font></p>";
            service.Update(email);
            SendEmail(service, email);
        }
        private Email GetEmail(IOrganizationService service, CodeActivityContext context)
        {
            Email myMail = (Email)service.Retrieve(Email.EntityLogicalName, DescEmail.Get(context).Id, new ColumnSet("description", "to"));
            return myMail;
        }
        private void SendEmail(IOrganizationService service, Email email)
        {
            Guid emailId = email.Id;
            SendEmailRequest sendEmailreq = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = "",
                IssueSend = true
            };
            SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailreq);
        }
    }
}
