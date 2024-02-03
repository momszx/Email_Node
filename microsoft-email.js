module.exports = function (RED) {
    function MicrosoftEmailNode(config) {
        RED.nodes.createNode(this, config);
        var node = this;
        const clientId = this.credentials.clientId;
        const clientSecret = this.credentials.clientSecret;
        const tenantId = this.credentials.tenantId;
        node.on('input', async function (msg) {
            try {
                /*var username = this.credentials.username;
                var password = this.credentials.password;
                node.warn(username);
                node.warn(password);*/

                // Get access token
                const accessToken = await getAccessToken(msg);

                // Send email using access token
                await sendEmail(accessToken, msg);

                console.log("Email sent successfully!");

                // Pass the modified message to the next node in the flow
                node.send(msg);
            } catch (error) {
                console.error("Error:", error.message);
                node.error("Error: " + error.message, msg);
            }
        });

        async function getAccessToken(input) {

            const scope = "https://graph.microsoft.com/.default";
            const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

            const payload = new URLSearchParams({
                client_id: clientId,
                scope: scope,
                client_secret: clientSecret,
                grant_type: "client_credentials"
            });

            const headers = {
                "Content-Type": "application/x-www-form-urlencoded"
            };

            const response = await fetch(url, {
                method: 'POST',
                body: payload,
                headers: headers,
            });
            
            const data = await response.json();
            return data.access_token;
        }

        async function sendEmail(accessToken, input) {
            const userEmail = input.payload.userEmail;
            const url = `https://graph.microsoft.com/v1.0/users/${userEmail}/sendMail`;

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

            await fetch(url, {
                method: 'POST',
                body: JSON.stringify(payload),
                headers: headers,
            });
        }
    }

    RED.nodes.registerType("microsoft-email", MicrosoftEmailNode,{
        credentials: {
            clientId: {type:"text"},
            tenantId: {type:"text"},
            clientSecret: {type:"password"}
        }
    });
};
