import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { TodoApp } from './todoApp';
import { TodoModel } from './core';
import { Authenticator } from '@microsoft/office-js-helpers';
import * as microsoftTeams from '@microsoft/teams-js';

var model = new TodoModel('react-todos');

function render() {
    ReactDOM.render(
        <TodoApp model={model} />,
        document.getElementsByClassName('todoapp')[0]
    );
}

/** initialize throws if called more than once and hence is wrapped in a try-catch to perform a safe initialization. */

try {
    microsoftTeams.initialize();
}
catch (e) {
}

if (!Authenticator.isAuthDialog(true /* Use Microsoft Teams Dialog */)) {
    model.subscribe(render);
    render();
}
