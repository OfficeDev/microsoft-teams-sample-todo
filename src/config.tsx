import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Authenticator } from './auth';
import { OutlookTasks } from './services';
import { IProfile } from './core';

class Config extends React.Component<any, any> {
    state: {
        loggedIn: boolean,
        profile: IProfile
    };
    outlookTasks: OutlookTasks;

    constructor(props: any) {
        super(props);
        this.state = {
            loggedIn: false,
            profile: null
        };

        this.outlookTasks = new OutlookTasks();
        var token = this.outlookTasks.graphToken;
        if (token) this.login();
    }

    render() {
        var content;
        if (this.state.loggedIn) {
            content = (
                <section className="ms-Persona ms-Persona--header">
                    <div className="ms-Persona ms-Persona--xl">
                        <div className="ms-Persona-imageArea">
                            <img className="ms-Persona-image"
                                title={this.state.profile.displayName}
                                alt={this.state.profile.displayName}
                                src={this.state.profile.thumbnail} />
                        </div>
                        <div className="ms-Persona-details">
                            <p className="ms-font-m">Signed in as</p>
                            <p className="ms-font-xxl">{this.state.profile.displayName}</p>
                            <p className="ms-font-m-plus">{this.state.profile.mail}</p>
                        </div>
                    </div>
                    <button className="ms-Button ms-Button--primary" onClick={e => this.logout()}><span className="ms-Button-label">Sign out</span></button>
                </section >
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
        return this.outlookTasks.login(true)
            .then(token => this.outlookTasks.profile())
            .then(profile => this.setState({
                loggedIn: true,
                profile: profile
            }))
            .then(token => this.completeSave())
            .catch(error => console.error('Failed to login using your Microsoft Account', error));
    }

    logout() {
        this.outlookTasks.logout();
        this.setState({
            loggedIn: false,
            profile: null
        });
    }

    completeSave() {
        if (!microsoftTeams) return;
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

if (!Authenticator.isAuthDialog()) {
    ReactDOM.render(
        <Config />,
        document.querySelector('.configApp')
    )
}