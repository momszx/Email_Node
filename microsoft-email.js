//const axios = require('axios');

module.exports = function (RED) {
    function MicrosoftEmailNode(config) {
        RED.nodes.createNode(this, config);
        var node = this;

        node.on('input', async function (msg) {
            try {
                node.warn("Belépési értéket");
                node.warn(msg.payload);
                // Get access token
                const accessToken = await getAccessToken(msg);

                // Send email using access token
                await sendEmail(accessToken,msg);

                console.log("Email sent successfully!");

                // Pass the modified message to the next node in the flow
                node.send(msg);
            } catch (error) {
                console.error("Error:", error.message);
                node.error("Error: " + error.message, msg);
            }
        });

        async function getAccessToken( input) {
            const clientId = input.payload.clientId;
            const clientSecret = input.payload.clientSecret;
            const tenantId = input.payload.tenantId;
            const scope = "https://graph.microsoft.com/.default";
            const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

            node.warn("Ide is bejutok");
            const payload = {
                client_id: clientId,
                scope: scope,
                client_secret: clientSecret,
                grant_type: "client_credentials"
            };

            const headers = {
                "Content-Type": "application/x-www-form-urlencoded"
            };

            const response = await axios.post(url, payload, { headers });
            return response.data.access_token;
        }

        async function sendEmail(accessToken,input) {
            const url = "https://graph.microsoft.com/v1.0/users/kalauz.mate@scepiconsulting.hu/sendMail";

            const payload = {
                message: {
                    subject: input.payload.subject,
                    body: {
                        contentType: "HTML",
                        content: input.payload.body
                    },
                    toRecipients: [
                        {
                            emailAddress: {
                                address: input.payload.address
                            }
                        }
                    ]
                },
                saveToSentItems: "true"
            };

            const headers = {
                "Authorization": "Bearer " + accessToken,
                "Content-Type": "application/json"
            };

            await axios.post(url, payload, { headers });
        }
    }

    RED.nodes.registerType("microsoft-email", MicrosoftEmailNode);
};
