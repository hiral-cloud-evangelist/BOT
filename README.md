## Azure Bot Framework with Multi-tenant Azure AD authentication

Azure BOT with Azure AD Sign in and LUIS capabilities that provides late coming member list and allow admin to auto deduct leave or send intimation via email as per action taken.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

At first step, BOT would represent SignIn Card. Login with Azure AD user account, on sucessfully login you can access BOT functionality.

# Code to show Signin card "BasicLuisDialog"

  #region Login

        public async Task<string> UserInfo(string Token)
        {
            HttpClient client = new HttpClient();

            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + Token);

            HttpResponseMessage response = await client.GetAsync(Constants.MicrosoftProfileUrl);

            string retResp = await response.Content.ReadAsStringAsync();
            AzureRespose data = JsonConvert.DeserializeObject<AzureRespose>(retResp);

            return data.EmailAddress;
        }

        private async Task LogIn(IDialogContext context)
        {
            string token;
            if (!context.PrivateConversationData.TryGetValue(AuthTokenKey, out token))
            {
                var conversationReference = context.Activity.ToConversationReference();

                context.PrivateConversationData.SetValue("persistedCookie", conversationReference);

                var reply = context.MakeMessage();
                reply.Type = "message";

                if (context.Activity.ChannelId == ChannelIds.Skype.ToString())
                {
                    Microsoft.Bot.Connector.Attachment plAttachment = GetSkypeSigninCard(conversationReference);
                    reply.Attachments.Add(plAttachment);
                }
                else
                {
                    Microsoft.Bot.Connector.Attachment plAttachment = GetSigninCard(conversationReference);
                    reply.Attachments.Add(plAttachment);
                }

                await context.PostAsync(reply);

                context.Wait(MessageReceivedAsync);
            }
            else
            {
                await context.PostAsync($"Your are already logged in.");
                context.Done(token);
            }
        }

        private static Microsoft.Bot.Connector.Attachment GetSkypeSigninCard(ConversationReference conversationReference)
        {
            var signinCard = new SigninCard
            {
                Text = "Please login to microsoft account",
                Buttons = new List<CardAction> { new CardAction(ActionTypes.Signin, "Authentication Required", value: SharepointHelpers.GetSharepointLoginURL(conversationReference, Constants.SharepointOauthCallback.ToString())) }
            };

            return signinCard.ToAttachment();
        }

        private static Microsoft.Bot.Connector.Attachment GetSigninCard(ConversationReference conversationReference)
        {
            List<CardAction> cardButtons = new List<CardAction>();
            CardAction plButton = new CardAction()
            {
                Value = SharepointHelpers.GetSharepointLoginURL(conversationReference, Constants.SharepointOauthCallback.ToString()),
                Type = "openUrl",
                Title = "Authentication Required"
            };
            cardButtons.Add(plButton);

            SigninCard plCard = new SigninCard("Please login to microsoft account", new List<CardAction>() { plButton });
            return plCard.ToAttachment();
        }
        #endregion
    }

2. On clicking Sign in, BOT would redirect user to sign in in azure, once authenticated user can use Time officer BOT. 

# authentication callback

    public class OAuthCallbackController : ApiController
    {
        [HttpGet]
        [Route("api/OAuthCallback")]
        public async Task<HttpResponseMessage> OAuthCallback([FromUri] string code, [FromUri] string session_state, string state, CancellationToken token)
        {

            var dict = HttpUtility.ParseQueryString(state);
            string json = JsonConvert.SerializeObject(dict.Cast<string>().ToDictionary(k => k, v => dict[v]));
            Address encodedAddress = JsonConvert.DeserializeObject<Address>(json);
            Address address = new Address(
                botId: SharepointHelpers.TokenDecoder(encodedAddress.BotId),
                channelId: SharepointHelpers.TokenDecoder(encodedAddress.ChannelId),
                conversationId: SharepointHelpers.TokenDecoder(encodedAddress.ConversationId),
                serviceUrl: SharepointHelpers.TokenDecoder(encodedAddress.ServiceUrl),
                userId: SharepointHelpers.TokenDecoder(encodedAddress.UserId)
                );

            var conversationReference = address.ToConversationReference();

            // Exchange the Sharepoint Auth code with Access token
            var accessToken = await SharepointHelpers.ExchangeCodeForAccessToken(conversationReference, code, Constants.SharepointOauthCallback.ToString());

            // Create the message that is send to conversation to resume the login flow
            var msg = conversationReference.GetPostToBotMessage();
            msg.Text = $"token:{accessToken}";

            // Resume the conversation to AuthDialog

            await Conversation.ResumeAsync(conversationReference, msg);

            using (var scope = DialogModule.BeginLifetimeScope(Conversation.Container, msg))
            {
                var dataBag = scope.Resolve<IBotData>();
                await dataBag.LoadAsync(token);
                ConversationReference pending;
                if (dataBag.PrivateConversationData.TryGetValue("persistedCookie", out pending))
                {
                    // remove persisted cookie
                    dataBag.PrivateConversationData.RemoveValue("persistedCookie");
                    await dataBag.FlushAsync(token);
                    return Request.CreateResponse("You are now logged in! Continue talking to the bot.");
                }
                else
                {
                    // Callback is called with no pending message as a result the login flow cannot be resumed.
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, new InvalidOperationException("Cannot resume!"));
                }
            }
        }

    }
    
3. SharePointHelpers contains code for authorization callback that can be reused for office 365 authentication in other application also. 
    
4. Replace Constants values as per needed.

### References to create Azure BOT Service, Developing, Testing, and publish on Azure.

1. Create a bot with Bot Service - https://docs.microsoft.com/en-us/azure/bot-service/bot-service-quickstart?view=azure-bot-service-3.0
2. Create a bot with the Bot Builder SDK for .NET - https://docs.microsoft.com/en-us/azure/bot-service/dotnet/bot-builder-dotnet-quickstart?view=azure-bot-service-3.0
3. Call a LUIS endpoint using C# - https://docs.microsoft.com/en-us/azure/cognitive-services/luis/luis-get-started-cs-get-intent


### Installing

A step by step series of examples that tell you how to get a development env running

1. Download / Clone code. 
2. Repleace web.config values as below. 

You will get MicrosoftAppId and MicrosoftAppPassword on creating azure bot app. You will get LUIS api key and App id on creating LUIS app. Copy those values and replace to below code.

```
    <add key="MicrosoftAppId" value="" />
    <add key="MicrosoftAppPassword" value="" />
    <add key="LuisAPIKey" value=""/>
    <add key="LuisAppId" value="" />
```

3. Create Azure Active Directory application https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-integrating-applications. Copy value of application id and application secret and replace values of SharepointAppId, and SharepointAppSecret respectively [ Constants.Cs class ]


## Debug Code

[Debug bots with the Bot Framework Emulator](https://docs.microsoft.com/en-us/azure/bot-service/bot-service-debug-emulator?view=azure-bot-service-3.0)

## Deployment

1. [Deploy your bot to Azure](https://docs.microsoft.com/en-us/azure/bot-service/bot-builder-howto-deploy-azure?view=azure-bot-service-3.0)
2. [Publish a bot to Bot Service](https://docs.microsoft.com/en-us/azure/bot-service/bot-service-continuous-deployment?view=azure-bot-service-3.0)


## Authors

* **Hiral Patel** (https://github.com/mehiralpatel)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details


    
5. In case, you want to more about code, or any difficulties please contact at mehiralpatel@gmail.com.
    
