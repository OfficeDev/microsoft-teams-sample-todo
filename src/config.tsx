import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Authenticator, IToken } from './auth';

import 'office-ui-fabric-core/dist/css/fabric.min.css';
import 'office-ui-fabric-js/dist/css/fabric.components.min.css';
import './app.css';

class Config extends React.Component<any, any> {
    state: any;
    authenticator: Authenticator;

    constructor(props: any) {
        super(props);

        this.authenticator = new Authenticator();
        this.authenticator.endpoints.registerMicrosoftAuth('73d044ea-4ae0-4bd8-b26c-7c2f924410a2', {
            scope: 'https://outlook.office.com/tasks.readwrite'
        });

        this.state = {
            loggedIn: false
        };
    }

    render() {
        var content;
        if (this.state.loggedIn) {
            content = (
                <section className="ms-Persona ms-Persona--header">
                    <div className="ms-Persona ms-Persona--xl">
                        <div className="ms-Persona-imageArea">
                            <img className="ms-Persona-image"
                                title="Bhargav Krishna"
                                alt="Bhargav Krishna"
                                src="https://avatars1.githubusercontent.com/u/3244360?v=3&s=466" />
                        </div>
                        <div className="ms-Persona-details">
                            <p className="ms-font-m">Signed in as</p>
                            <p className="ms-font-xxl">Bhargav Krishna</p>
                            <p className="ms-font-m-plus">bhargav.krishna@microsoft.com</p>
                        </div>
                    </div>
                </section>
            )
        }
        else {
            content = (
                <button className="microsoft-identity" onClick={e => this.login()}>
                    <span className="microsoft-identity-logo"></span>
                    <span className="microsoft-identity-label">Sign in with Microsoft</span>
                </button >
            )
        }

        return (
            <div>{content}</div>
        )
    }

    login() {
        return this.authenticator.authenticate('Microsoft')
            .then(token => this.setState({ loggedIn: true }))
            .then(token => this.completeSave())
            .catch(error => console.error('Failed to login using your Microsoft Account', error));
    }

    completeSave() {
        if (!microsoftTeams) return;
        microsoftTeams.initialize();
        microsoftTeams.settings.setValidityState(true);
        microsoftTeams.settings.registerOnSaveHandler(saveEvent => {
            microsoftTeams.settings.setSettings({
                contentUrl: "https://localhost:3000/index.html",
                suggestedDisplayName: "My Tasks",
                websiteUrl: "https://localhost:3000/index.html"
            });
            saveEvent.notifySuccess();
        });
    }
}

ReactDOM.render(
    <Config />,
    document.querySelector('.configApp')
)

