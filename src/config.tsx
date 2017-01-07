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
            /* Initialize the Teams library before any other SDK calls.
             * Initialize throws if called more than once and hence is wrapped in a try-catch to perform a safe initialization.
             */

            microsoftTeams.initialize();
        }
        catch (e) {
        }
        finally {
            this.state = {
                groupId: null,
                upn: null
            };

            /** Pass the Context interface to the initialize function below */
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

    /**
     * This will save off the Context and registers the Save Handler, which will be called when the user attempts to Save.  This handler will be used to save the Settings for the Tab content.
     * @param {object} groupId - The O365 group id for the team.
     * @param {object} upn - The current user's upn (user principal name).
     */
    initialize({ groupId, upn}) {
        this.setState({ groupId, upn });
        console.log(this.state);
        /** Enable the Save button  */
        microsoftTeams.settings.setValidityState(true);
        /** Register the save handler */
        microsoftTeams.settings.registerOnSaveHandler(saveEvent => {
            /** Store Tab content settings */
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