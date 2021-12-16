import React from 'react';
import AuthService from '../services/AuthService'
import { List, Icon } from "@fluentui/react";
import * as MicrosoftGraphClient from "@microsoft/microsoft-graph-client";

/**
 * The web UI used when Teams pops out a browser window
 */
class Web extends React.Component {

    constructor(props) {
        super(props);

        this.state = {
            accessToken: null,
            messages: []
        }
    }

    componentWillMount() {
        console.log('Provider', List.Item);
        if (!AuthService.isLoggedIn()) {
            // Will redirect the browser and not return; will redirect back if successful
            AuthService.login(["User.Read", "Mail.Read"]);
        } else {
            this.msGraphClient = MicrosoftGraphClient.Client.init({
                authProvider: async (done) => {
                    if (!this.state.accessToken) {
                        // Might redirect the browser and not return; will redirect back if successful
                        const token = await AuthService.getAccessToken(["User.Read", "Mail.Read"]);

                        this.setState({
                            accessToken: token
                        });
                    }

                    done(null, this.state.accessToken);
                }
            });
        }
    }

    render() {
        return (
            <div>
                <h1>MSAL 2.0 Test App</h1>
                <button onClick={this.getMessages.bind(this)}>Get Mail</button>
                <p>Username: {AuthService.getUsername()}</p>
                <ul>
                {
                    this.state.messages.map(message => (
                        <li key={message.id}>
                            {message.subject}
                            {/* key={message.id}
                            header={message.receivedDateTime}
                            content={message.subject}> */}
                        </li>
                    ))
                }
                </ul>
            </div>
        );
    }

    getMessages() {
        this.msGraphClient
            .api("me/mailFolders/inbox/messages")
            .select(["receivedDateTime", "subject"])
            .top(15)
            .get(async (error, rawMessages, rawResponse) => {
                if (!error) {
                    this.setState(Object.assign({}, this.state, {
                        messages: rawMessages.value
                    }));
                } else {
                    this.setState({
                        error: error
                    });
                }
            });
    }
}

export default Web;