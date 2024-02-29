using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Email")]
    public class EmailSteps
    {
        public async void SendEmail(/*SendMailPostRequestBody messageBody*/)
        {
            try
            {
                //return Exchange365Auth.GraphClient.Me.SendMail.PostAsync(messageBody).Status.ToString();
                
                var requestBody = new SendMailPostRequestBody
                {
                    Message = new Message
                    {
                        Subject = "Hey Jerry",
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Text,
                            Content = "These pretzels are making me thirsty!",
                        },
                        ToRecipients = new List<Recipient>
                        {
                            new Recipient
                            {
                                EmailAddress = new EmailAddress
                                {
                                    Address = "shawn.lirette@decisions.com",
                                },
                            },
                        },
                        CcRecipients = null,
                    },
                    SaveToSentItems = false,
                };
                
                await Exchange365Auth.GraphClient.Me.SendMail.PostAsync(requestBody);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
    }
}