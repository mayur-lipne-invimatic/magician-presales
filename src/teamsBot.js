// Required libraries and modules
const {
    TeamsActivityHandler,
    MessageFactory,
    InputHints,
    CardFactory,
    TeamsInfo
} = require('botbuilder');
const axios = require('axios');

// Constants
const BACKEND_URL = `https://api.${process.env.LOOPS_HOST}/conversation/v1/webhooks/msteams/events`;

// TeamsBot class that handles interactions within MS Teams
class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();


        // Event handler for when a new message is received.
        this.onMessage(async (context, next) => {
            try {

                // Guard: let TeamsFx ConversationBot handle 'loops'/'ask' commands; skip generic handlers to avoid duplicates.
                try {
                    const raw = (context && context.activity && context.activity.text) ? String(context.activity.text) : '';
                    const text = raw.trim();
                    if (/^(loops|ask)\b/i.test(text)) {
                        await next();
                        return; // do not continue to generic notifier paths
                    }
                } catch (e) {
                    console.error('command guard failed', e);
                }
                
                if (context.activity) {
                    // Extract Adaptive Card submit data
                    const formData = context.activity.value;

                    if (context.activity && context.activity.text) {
                        const text = context.activity.text.trim().toLowerCase();
                        console.log("Received message:", text);

                        // Check if the message is in a personal chat
                        if (context.activity.conversation.conversationType === "personal") {
                            // Notify the backend that a new message was received in a direct chat with the bot.
                            await notifyBackendForNewMessage('UserInteractionWithBot', context);
                        } else {
                            // Check if the message mentions the bot.
                            const isBotMentioned = context.activity.entities.some(entity => entity.type === 'mention' && entity.mentioned.id === context.activity.recipient.id);

                            // Determine the appropriate action based on the message content.
                            if (isBotMentioned) {
                                const words = text.split(' ');
                                const loopsIndex = words.indexOf("loops");

                                if (loopsIndex === -1) {
                                    // Notify the backend that a new message was received.
                                    const actionType = 'LoopsBotMentioned';
                                    await notifyBackendForNewMessage(actionType, context);
                                }
                            } else {
                                // If the bot is not mentioned, still notify the backend about the new message.
                                const conversationId = context.activity.conversation.id;
                                const activityId = context.activity.id;

                                const messageIdMatch = conversationId.match(/messageid=(\d+)/);
                                const actionType = ((messageIdMatch && messageIdMatch[1]) && messageIdMatch[1] !== activityId) ? 'NewThreadedMessage' : 'NewMessage';
                                await notifyBackendForNewMessage(actionType, context);
                            }
                        }
                        await next();
                    } else if (formData) {
                        try {
                            if (formData.dropdownChanged) {
                                const actionType = 'UpdateAdaptiveCardOnDropDown';
                                const dropdownValue = cardInput.dropdown;

                                // Create the payload
                                const payload = {
                                    ...await constructPayloadForBackend(actionType, formData, context),
                                    selectedValue: dropdownValue
                                };

                                // Update the card based on the dropdownValue
                                const updatedCard = await generateUpdatedCard(payload);
                                await context.updateActivity({
                                    attachments: [updatedCard]
                                });
                            } else {
                                const actionType = 'ProcessAdaptiveCardForm';
                                // Create the payload
                                const payload = {
                                    ...await constructPayloadForBackend(actionType, formData, context),
                                    from: context.activity.from,
                                    to: context.activity.recipient,
                                    threadId: extractThreadId(context.activity.conversation.id),
                                    eventTimestamp: context.activity.timestamp,
                                    eventId: context.activity.id,
                                    hiddenData: context.activity.value.hiddenData
                                };

                                // Default 'isButton' and 'responseAsReply' to false if not present
                                // Extract buttonData
                                const buttonData = formData.buttonData || {};

                                // Default 'isButton' and 'responseAsReply' to false if not present
                                const isButton = buttonData.isButton || false;
                                const responseAsReply = buttonData.responseAsReply || false;

                                const payloadJsonString = JSON.stringify(payload);
                                const response = await axios.post(BACKEND_URL, payloadJsonString);
                                await processBackendResponse(response, context, responseAsReply);
                            }
                        } catch (error) {
                            console.error('Error in onMessage:', error);
                            await context.updateActivity(error.message);
                        }
                    }
                }

            } catch (error) {
                console.error('Error in onMessage:', error);
                await context.sendActivity(error.message);
            }
        });



        // Event handler for when a new channel is created.
        this.onConversationUpdate(async (context, next) => {
            try {

                if (context.activity.channelData.eventType == "teamMemberAdded") {
                    await handleTeamMembersAdded(context);
                } else if (context.activity.channelData.eventType == "teamMemberRemoved") {
                    await handleTeamMembersRemoved(context);
                }

                await next();
            } catch (error) {
                console.error('Error in onConversationUpdate:', error);
            }
        });

        this.onInvoke = async (context, next) => {
            try {
                // Check if the invoke activity is from an Adaptive Card action
                if (context.activity.name === 'adaptiveCard/action') {
                    const invokeValue = context.activity.value;

                    // Process the action data (e.g., form inputs from the Adaptive Card)
                    const actionData = invokeValue.action.data;

                    // You can handle different types of actions here
                    switch (invokeValue.action.verb) {
                        case 'submit':
                            // Process the submit action
                            await this.handleCardSubmitAction(context, actionData);
                            break;
                        // Add more cases for other action verbs if needed
                        default:
                            // Handle unknown action verbs
                            await context.sendActivity(`Unknown action verb: ${invokeValue.action.verb}`);
                            break;
                    }

                    // Send a response back to the Adaptive Card
                    await context.sendActivity({
                        type: 'invokeResponse',
                        value: {
                            status: 200,
                            body: {
                                type: 'application/vnd.microsoft.card.adaptive',
                                content: {
                                    type: 'AdaptiveCard',
                                    body: [{
                                        type: 'TextBlock',
                                        text: 'Action received!'
                                    }],
                                    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                                    version: '1.3'
                                }
                            }
                        }
                    });
                }

                await next();
            } catch (error) {
                console.error('Error in onInvoke:', error);
                // Send a failure response back
                await context.sendActivity({
                    type: 'invokeResponse',
                    value: {
                        status: 500,
                        body: {
                            error: error.message
                        }
                    }
                });
            }
        };

    }

    // Function to handle the card submit action
    async handleCardSubmitAction(context, actionData) {
        // Process actionData here (e.g., save to database, call an API, etc.)
        // For now, let's log the data
        console.log('Received submit action data:', actionData);

        // Optionally send a response message to the user
        await context.sendActivity('Your form data was received and processed.');
    }

}



/**
 * Notifies the backend with a new message by sending a POST request.
 */
async function notifyBackendForNewMessage(actionType, context) {

    // Check if the bot is mentioned
    const isBotMentioned = context.activity.entities && context.activity.entities.some(entity => entity.type === 'mention' && entity.mentioned.id === context.activity.recipient.id);

    // Construct the payload to be sent to the backend.
    const payload = {
        ...await constructPayloadForBackend(actionType, null, context),
        messageText: context.activity.text,
        eventId: context.activity.id,
        threadId: extractThreadId(context.activity.conversation.id),
        eventTimestamp: context.activity.timestamp,
        from: context.activity.from,
        to: context.activity.recipient,
        isBotMentioned: isBotMentioned
    };
    const payloadJsonString = JSON.stringify(payload);

    try {
        // Sending a POST request to the backend with the constructed payload.
        await axios.post(BACKEND_URL, payloadJsonString, {
            headers: {
                'Content-Type': 'application/json'
            }
        });
    } catch (error) {
        // Log the error if the backend notification fails.
        console.error(`Error notifying backend for ${actionType}:`, error);
    }
}



/**
 * Constructs a payload to be sent to the backend containing
 */
async function constructPayloadForBackend(actionType, data, context) {
    // Extract channel details from the context.
    const channelId = context.activity.channelData && context.activity.channelData.channel ? context.activity.channelData.channel.id : 'unknown';
    let teamId = context.activity.channelData.team ? context.activity.channelData.team.id : undefined;
    let teamName;

    if (teamId) {
        try {
            // Fetch team details using the teamId.
            const teamDetails = await TeamsInfo.getTeamDetails(context, teamId);
            teamName = teamDetails.name;
            teamId = teamDetails.aadGroupId;
        } catch (err) {
            console.error("Error fetching team details:", err);
            teamName = "Unknown"; // Or any other default value.
        }
    } else {
        teamName = "Unknown"; // Default value if no team data is available.
    }

    // Return the constructed payload.
    return {
        action: actionType,
        conversationId: channelId,
        channelId: channelId,
        teamsTenantId: context.activity.channelData.tenant.id,
        teamId: teamId,
        teamName: teamName,
        data: data,
        rawEvent: context
    };
}



/**
 * Event handler for when a user joins the team.
 */
async function handleTeamMembersAdded(context) {
    try {
        const membersAdded = context.activity.membersAdded;

        // Iterate through all members that have been added.
        for (let member of membersAdded) {

            // Check if the bot itself is the member that was added.
            if (member.id === context.activity.recipient.id) {

                // Extract relevant details about the bot and the team.
                const botName = context.activity.recipient.name;
                let teamName = context.activity.channelData.team ? context.activity.channelData.team.name : "Unknown";

                // Define an action to fetch the welcome message for the bot.
                let actionType = 'FetchWelcomeMessage';
                let teamId = context.activity.channelData.team ? context.activity.channelData.team.id : undefined;

                if (teamId) {
                    try {
                        // Fetch team details using the teamId.
                        const teamDetails = await TeamsInfo.getTeamDetails(context, teamId);
                        teamName = teamDetails.name;
                        teamId = teamDetails.aadGroupId;
                    } catch (err) {
                        console.error("Error fetching team details:", err);
                        teamName = "Unknown"; // Or any other default value.
                    }
                } else {
                    teamName = "Unknown"; // Default value if no team data is available.
                }
                const payload = {
                    action: actionType,
                    teamsTenantId: context.activity.channelData.tenant.id,
                    teamId: teamId,
                    teamName: teamName,
                    from: context.activity.from,
                    to: context.activity.recipient,
                    rawEvent: context
                };

                const payloadJsonString = JSON.stringify(payload);

                // Attempt to fetch the welcome message for the bot from the backend.
                try {
                    const welcomeResponse = await axios.post(BACKEND_URL, payloadJsonString, {
                        headers: {
                            'Content-Type': 'application/json'
                        }
                    });

                    // Extract the welcome message or use a default message if none is provided.
                    let welcomeMessage = welcomeResponse.data && welcomeResponse.data.message ?
                        welcomeResponse.data.message :
                        `<at>${botName}</at> was added to ${teamName} by an administrator`;

                    // Craft a message with mentions and send it.
                    const message = MessageFactory.text(welcomeMessage, welcomeMessage, InputHints.ExpectingInput);
                    message.entities = [{
                        mentioned: context.activity.recipient,
                        text: `<at>${botName}</at>`,
                        type: 'mention'
                    }];
                    await context.sendActivity(message);

                    // Notify the backend about the addition of the bot to a team or channel.
                    actionType = 'NotifyBotAddition';
                    let teamId = context.activity.channelData.team ? context.activity.channelData.team.aadGroupId : undefined;

                    if (teamId) {
                        try {
                            // Fetch team details using the teamId.
                            const teamDetails = await TeamsInfo.getTeamDetails(context, teamId);
                            teamName = teamDetails.name;
                            teamId = teamDetails.aadGroupId;
                        } catch (err) {
                            console.error("Error fetching team details:", err);
                            teamName = "Unknown"; // Or any other default value.
                        }
                    } else {
                        teamName = "Unknown"; // Default value if no team data is available.
                    }
                    const notifyPayload = {
                        action: actionType,
                        teamsTenantId: context.activity.channelData.tenant.id,
                        teamId: teamId,
                        teamName: teamName,
                        channelId: context.activity.channelData.channel.id,
                        from: context.activity.from,
                        to: context.activity.recipient,
                        rawEvent: context
                    };

                    const notifyPayloadJsonString = JSON.stringify(notifyPayload);
                    await axios.post(BACKEND_URL, notifyPayloadJsonString, {
                        headers: {
                            'Content-Type': 'application/json'
                        }
                    });

                } catch (error) {
                    // If there's an error fetching the welcome message, log it and send a default message.
                    console.error("Error fetching welcome message from backend:", error);
                    if (defaultMessage) {
                        await context.sendActivity(defaultMessage);
                    }
                }
            } else {
                // For members other than the bot.
                const actionType = 'MemberJoinedChannel';
                let teamId = context.activity.channelData.team ? context.activity.channelData.team.aadGroupId : undefined;

                if (teamId) {
                    try {
                        // Fetch team details using the teamId.
                        const teamDetails = await TeamsInfo.getTeamDetails(context, teamId);
                        teamName = teamDetails.name;
                        teamId = teamDetails.aadGroupId;
                    } catch (err) {
                        console.error("Error fetching team details:", err);
                        teamName = "Unknown"; // Or any other default value.
                    }
                } else {
                    teamName = "Unknown"; // Default value if no team data is available.
                }
                const payload = {
                    action: actionType,
                    teamsTenantId: context.activity.channelData.tenant.id,
                    teamId: teamId,
                    teamName: teamName,
                    userId: member.aadObjectId,
                    from: context.activity.from,
                    to: context.activity.recipient,
                    rawEvent: context
                };

                const payloadJsonString = JSON.stringify(payload);

                // Notify the backend about the addition of a member to a team or channel.
                await axios.post(BACKEND_URL, payloadJsonString, {
                    headers: {
                        'Content-Type': 'application/json'
                    }
                });
            }
        }
    } catch (error) {
        console.error('Error in onMembersAdded:', error);
    }
}



/**
 * Event handler for when a user is removed from the team.
 */
async function handleTeamMembersRemoved(context) {
    try {
        const membersRemoved = context.activity.membersRemoved;

        // Iterate through all members that have been removed.
        for (let member of membersRemoved) {

            // Check if the bot itself is the member that was removed.
            if (member.id === context.activity.recipient.id) {
                const botName = context.activity.recipient.name;
                const teamName = context.activity.channelData.team ? context.activity.channelData.team.name : "a channel";
                const removerName = context.activity.from.name;

                // Craft a message indicating the bot was removed and by whom, then send it.
                const removalMessage = `<at>${botName}</at> was removed from ${teamName} by <at>${removerName}</at>.`;
                await context.sendActivity(removalMessage);

                // Notify the backend about the removal of the bot to a team or channel.
                actionType = 'NotifyBotRemoval';
                let teamId = context.activity.channelData.team ? context.activity.channelData.team.aadGroupId : undefined;

                if (teamId) {
                    try {
                        // Fetch team details using the teamId.
                        const teamDetails = await TeamsInfo.getTeamDetails(context, teamId);
                        teamName = teamDetails.name;
                        teamId = teamDetails.aadGroupId;
                    } catch (err) {
                        console.error("Error fetching team details:", err);
                        teamName = "Unknown"; // Or any other default value.
                    }
                } else {
                    teamName = "Unknown"; // Default value if no team data is available.
                }
                const notifyPayload = {
                    action: actionType,
                    teamsTenantId: context.activity.channelData.tenant.id,
                    teamId: teamId,
                    teamName: teamName,
                    from: context.activity.from,
                    to: context.activity.recipient,
                    rawEvent: context
                };

                const notifyPayloadJsonString = JSON.stringify(notifyPayload);
                await axios.post(BACKEND_URL, notifyPayloadJsonString, {
                    headers: {
                        'Content-Type': 'application/json'
                    }
});
            } else {
                // For members other than the bot.
                const actionType = 'MemberLeftChannel';
                let teamId = context.activity.channelData.team ? context.activity.channelData.team.id : undefined;

                if (teamId) {
                    try {
                        // Fetch team details using the teamId.
                        const teamDetails = await TeamsInfo.getTeamDetails(context, teamId);
                        teamName = teamDetails.name;
                        teamId = teamDetails.aadGroupId;
                    } catch (err) {
                        console.error("Error fetching team details:", err);
                        teamName = "Unknown"; // Or any other default value.
                    }
                } else {
                    teamName = "Unknown"; // Default value if no team data is available.
                }
                const payload = {
                    action: actionType,
                    teamsTenantId: context.activity.channelData.tenant.id,
                    teamId: teamId,
                    teamName: teamName,
                    userId: member.aadObjectId,
                    from: context.activity.from,
                    to: context.activity.recipient,
                    rawEvent: context
                };

                const payloadJsonString = JSON.stringify(payload);

                // Notify the backend about the removal of a member from a team or channel.
                await axios.post(BACKEND_URL, payloadJsonString, {
                    headers: {
                        'Content-Type': 'application/json'
                    }
                });
            }
        }
    } catch (error) {
        console.error('Error in onMembersRemoved:', error);
    }
}



/**
 * Event handler for when a channel is created for the team.
 */
async function handleTeamsChannelCreated(context) {
    try {
        const channelId = context.activity.channelData && context.activity.channelData.channel ? context.activity.channelData.channel.id : 'unknown';
        const channelName =  context.activity.channelData && context.activity.channelData.channel ? context.activity.channelData.channel.name : 'unknown';
        let teamId = context.activity.channelData && context.activity.channelData.team ? context.activity.channelData.team.id : undefined;

        if (teamId) {
            try {
                // Fetch team details using the teamId.
                const teamDetails = await TeamsInfo.getTeamDetails(context, teamId);
                teamName = teamDetails.name;
                teamId = teamDetails.aadGroupId;
            } catch (err) {
                console.error("Error fetching team details:", err);
                teamName = "Unknown"; // Or any other default value.
            }
        } else {
            teamName = "Unknown"; // Default value if no team data is available.
        }

        // Create the payload to send to the backend.
        const actionType = 'NotifyChannelCreation';
        const payload = {
            action: actionType,
            teamsTenantId: context.activity.channelData.tenant.id,
            teamId: teamId, // Here you might want to use teamId directly, depending on your needs
            channelId: channelId,
            channelName: channelName,
            teamName: teamName,
            from: context.activity.from,
            to: context.activity.recipient,
            rawEvent: context
        };

        const payloadJsonString = JSON.stringify(payload);

        // Send the payload to the backend.
        await axios.post(BACKEND_URL, payloadJsonString, {
            headers: {
                'Content-Type': 'application/json'
            }
        });
    } catch (error) {
        console.error('Error in onTeamsChannelCreated:', error);
    }
}




/**
 * Process the response from the backend and send the appropriate message/card
 * to the user based on the response content.
 */
async function processBackendResponse(response, context, responseAsReply) {
    try {
        // Check if backend response indicates successful processing.
        if (response.data.status === "ok") {
            console.log("Backend response:", response.data);

            // If there's card JSON data in the response, create an adaptive card.
            if (response.data && response.data.cardJson) {
                const card = CardFactory.adaptiveCard(JSON.parse(response.data.cardJson));

                // If action type is not for fetching input, display card.
                if (response.data.actionType !== 'FetchInputAdaptiveCard') {
                    if (responseAsReply) {
                        await context.sendActivity({
                            type: 'message',
                            id: context.activity.replyToId,
                            attachments: [card]
                        });
                    } else {
                        await context.updateActivity({
                            type: 'message',
                            id: context.activity.replyToId,
                            attachments: [card]
                        });
                    }
                }
            }
            // If there's a message in the response, send it.
            else if (response.data && response.data.message) {
                if (responseAsReply) {
                    await context.sendActivity({
                        type: 'message',
                        id: context.activity.replyToId,
                        text: response.data.message
                    });
                } else {
                    await context.updateActivity({
                        type: 'message',
                        id: context.activity.replyToId,
                        text: response.data.message
                    });
                }
            }
        }
        // If backend response indicates an error, log it and inform the user.
        else {
            const defaultCardJson = getDefaultAdaptiveCard(response.data.cardJson);
            const card = CardFactory.adaptiveCard(defaultCardJson);
            if (responseAsReply) {
                await context.sendActivity({
                    type: 'message',
                    id: context.activity.replyToId,
                    attachments: [card],
                    channelData: {
                        notification: {
                            alert: true
                        },
                        microsoftTeams: {
                            entityID: context.activity.replyToId,
                            task: {
                                fetchTask: true,
                                type: "continue",
                                title: "Error",
                                height: 500,
                                width: 400,
                                value: {
                                    processingActivityId: context.activity.replyToId // Storing the processingActivity.id in the value
                                }
                            }
                        }
                    }
                });
            } else {
                await context.updateActivity({
                    type: 'message',
                    id: context.activity.replyToId,
                    attachments: [card],
                    channelData: {
                        notification: {
                            alert: true
                        },
                        microsoftTeams: {
                            entityID: context.activity.replyToId,
                            task: {
                                fetchTask: true,
                                type: "continue",
                                title: "Error",
                                height: 500,
                                width: 400,
                                value: {
                                    processingActivityId: context.activity.replyToId // Storing the processingActivity.id in the value
                                }
                            }
                        }
                    }
                });
            }
        }
    } catch (error) {
        console.error('Error processing the backend response:', error);
        await context.updateActivity({
            type: 'message', // Activity type
            id: context.activity.replyToId, // ID of the activity to be updated
            text: 'An unexpected error occurred. Please try again later.' // New text to replace the previous message
        });
    }
}



/**
 * Generate an updated adaptive card by sending the given payload to the backend.
 * If the backend fails to provide an updated card, generate a default error card.
 */
async function generateUpdatedCard(payload) {
    // Send the payload to your backend and wait for a response.
    const payloadJsonString = JSON.stringify(payload);
    const response = await axios.post(YOUR_BACKEND_URL, payloadJsonString);

    // If the backend response is successful and contains card data, return the card.
    if (response.data.status === "ok" && response.data.cardJson) {
        return CardFactory.adaptiveCard(response.data.cardJson);
    } else {
        // If the backend response indicates an error or doesn't provide card data,
        // return a default error card.
        return CardFactory.adaptiveCard({
            type: "AdaptiveCard",
            body: [{
                "type": "TextBlock",
                "text": "An error occurred while updating the card."
            }]
        });
    }
}



/**
 * The function uses a regular expression to identify and extract
 * a messageId from a provided string. The messageId is assumed to be
 * a series of one or more numeric characters following the string "messageid=".
 */
function extractThreadId(inputString) {
    // Define a regular expression pattern to identify and extract the messageId.
    // The pattern looks for the string "messageid=", followed by one or more numeric characters.
    // The parentheses () create a capturing group, so that the numeric part can be extracted directly.
    const messageIdRegex = /messageid=([0-9]+)/;

    // Apply the regular expression to the input string.
    // If the pattern is found, match will be an array where match[1] is the first captured group (the numeric part).
    // If the pattern is not found, match will be null.
    const match = inputString.match(messageIdRegex);

    // Check if the pattern was found in the input string.
    if (match && match[1]) {
        // If a messageId was found, return it.
        return match[1];
    } else {
        // If a messageId was not found, return null.
        return null;
    }
}



/**
 * Generates and returns the default Adaptive Card JSON structure.
 * This card is meant to be used as a fallback when the system is unable
 * to retrieve or render the primary form/data.
 */
function getDefaultAdaptiveCard(errorMessage) {
    return {
        // Specifies the type of the card.
        type: "AdaptiveCard",

        // Defines the version of the Adaptive Card schema that this card requires.
        version: "1.3",

        // Contains the primary card elements that compose the content of the card.
        body: [{
            // Text block element to show a bold title.
            type: "TextBlock",
            size: "Medium",
            weight: "Bolder",
            text: "Error"
        },
            {
                // Text block element to display the error message.
                type: "TextBlock",
                text: errorMessage,
                wrap: true // Allow the text to wrap and extend beyond one line.
            }
        ],

        // Specifies the schema this card uses. It's a standard practice to include this.
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json"
    };
}



module.exports.TeamsBot = TeamsBot;