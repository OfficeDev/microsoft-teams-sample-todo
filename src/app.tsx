import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { TodoApp } from './todoApp';
import { TodoModel } from './core';
import { Authenticator } from '@microsoft/office-js-helpers';

var model = new TodoModel('react-todos');

function render() {
    ReactDOM.render(
        <TodoApp model={model} />,
        document.getElementsByClassName('todoapp')[0]
    );
}

microsoftTeams.initialize();
if (!Authenticator.isTeamsDialog()) {
    model.subscribe(render);
    render();
}