# BOT

1.	At first step, when you connect to Time officer Bot in skype channel, BOT would represent SignIn Card. 

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
    
5. In case, you want to more about code, or any difficulties please contact at mehiralpatel@gmail.com.
    
