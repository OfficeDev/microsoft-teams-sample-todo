import * as React from 'react';
import * as ReactDOM from 'react-dom';

import 'todomvc-common/base.css';
import 'todomvc-app-css/index.css';
import 'todomvc-common/base.js';
import 'office-ui-fabric-core/dist/css/fabric.min.css';
import './app.css';

import { TodoApp } from './todoApp';
import { TodoModel } from './core';
import { OutlookTasks } from './services';

var model = new TodoModel('react-todos');

function render() {
    ReactDOM.render(
        <TodoApp model={model} />,
        document.getElementsByClassName('todoapp')[0]
    );

    var outlookTasks = new OutlookTasks();
    outlookTasks.get().then(tasks => console.log(tasks));
}

model.subscribe(render);
render();