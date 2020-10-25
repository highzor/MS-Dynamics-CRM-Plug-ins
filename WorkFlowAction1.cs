using System.Linq;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Crm.Sdk.Samples.EarlyBoundTypes;
using Microsoft.Xrm.Sdk.Query;

namespace Custom.WorkFlowAction1
{
    public class ChangeEmail : CodeActivity
    {
        [RequiredArgument]
        [Input("Email input")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> DescEmail { get; set; }

        [Output("String output")]
        public OutArgument<string> ChDescEmail { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);
            Email myMail = GetEmail(service, context);
            var partyInf = myMail.To.First();
            myMail.Description = $"<p><font size='5' color='red' face='Arial'>Ваша компания зарегистрирована {partyInf.PartyId.Name}. Будем рады продолжить сотрудничество.</font></p>";
            service.Update(myMail);
        }
        private Email GetEmail(IOrganizationService service, CodeActivityContext context)
        {
            Email myMail = (Email)service.Retrieve(Email.EntityLogicalName, DescEmail.Get(context).Id, new ColumnSet("description", "to"));
            return myMail;
        }
    }
}
