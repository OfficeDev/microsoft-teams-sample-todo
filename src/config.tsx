import * as React from 'react';
import { render } from 'react-dom';
import { Utilities } from '@microsoft/office-js-helpers';

interface IConfigState {
    groupId: string;
    upn: string;
}

class Config extends React.Component<any, any> {
    state: IConfigState;

    constructor(props: any) {
        super(props);
        if (!microsoftTeams) return;

        try {
            microsoftTeams.initialize();
            this.state = {
                groupId: null,
                upn: null
            };
        }
        catch (e) {
            Utilities.log(e);
        }
        finally {
            microsoftTeams.getContext(context => this.initialize(context as any));
        }
    }

    render() {
        let output;

        if (this.state) {
            output = (
                <section className="ms-Persona ms-Persona--header">
                    <div className="ms-Persona ms-Persona--xl">
                        <div className="ms-Persona-details">
                            <p className="ms-font-m">Authorized as</p>
                            <p className="ms-font-xxl">{this.state.upn}</p>
                            <p className="ms-font-m-plus">{this.state.groupId}</p>
                        </div>
                    </div>
                </section>
            );
        }
        else {
            output = (
                <section className="ms-Persona ms-Persona--header">
                    <div className="ms-Persona ms-Persona--xl">
                        <div className="ms-Persona-details">
                            <p className="ms-font-xxl">Welcome Admin</p>
                            <p className="ms-font-m-plus">Authorizing...</p>
                        </div>
                    </div>
                </section>
            )
        }

        return (
            <div>{output}</div>
        )
    }

    initialize({ groupId, upn}) {
        this.setState({ groupId, upn });
        console.log(this.state);
        microsoftTeams.settings.setValidityState(true);
        microsoftTeams.settings.registerOnSaveHandler(saveEvent => {
            microsoftTeams.settings.setSettings({
                contentUrl: `${location.origin}/index.html`,
                suggestedDisplayName: "My Tasks",
                websiteUrl: `${location.origin}/index.html`
            });
            saveEvent.notifySuccess();
        });
    }
}

render(
    <Config />,
    document.querySelector('.configApp')
)