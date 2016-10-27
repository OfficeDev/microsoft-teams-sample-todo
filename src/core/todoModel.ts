import { Utils } from './utils';
import { ITodo, ITodoModel } from './interfaces';
import { OutlookTasks } from '../services';

// Generic 'model' object. You can use whatever
// framework you want. For this application it
// may not even be worth separating this logic
// out, but we do this to demonstrate one way to
// separate out parts of your application.
export class TodoModel implements ITodoModel {

    public key: string;
    public todos: Array<ITodo>;
    public onChanges: Array<any>;
    private _outlookTasks: OutlookTasks;

    constructor(key) {
        this.key = key;
        this._outlookTasks = new OutlookTasks('https://outlook.office.com/tasks.readwrite');
        this.todos = [];
        this._outlookTasks.get().then(tasks => {
            this.todos = tasks;
            this.inform();
        })
        this.onChanges = [];
    }

    public subscribe(onChange) {
        this.onChanges.push(onChange);
    }

    public inform() {
        this.onChanges.forEach(function (cb) { cb(); });
    }

    public addTodo(title: string) {
        let todo: ITodo = {
            id: Utils.uuid(),
            title: title,
            completed: false
        };

        this._outlookTasks.create(todo).then(todo => {
            this.todos = this.todos.concat(todo);
            this.inform();
        });
    }

    public toggleAll(checked: Boolean) {
        // Note: It's usually better to use immutable data structures since they're
        // easier to reason about and React works very well with them. That's why
        // we use map(), filter() and reduce() everywhere instead of mutating the
        // array or todo items themselves.
        this.todos = this.todos.map<ITodo>((todo: ITodo) => {
            var update = Utils.extend({}, todo, { completed: checked });
            this._outlookTasks.markAsComplete(update);
            return update
        });

        this.inform();
    }

    public toggle(todoToToggle: ITodo) {
        this.todos = this.todos.map<ITodo>((todo: ITodo) => {
            return todo !== todoToToggle ?
                todo : (() => {
                    this._outlookTasks.markAsComplete(todo);
                    return Utils.extend({}, todo, { completed: !todo.completed });
                })();
        });

        this.inform();
    }

    public destroy(todo: ITodo) {
        this._outlookTasks.delete(todo).then(result => {
            if (result) {
                this.todos = this.todos.filter(function (candidate) {
                    return candidate !== todo;
                });
            }

            this.inform();
        });
    }

    public save(todoToSave: ITodo, text: string) {
        this.todos = this.todos.map(function (todo) {
            return todo !== todoToSave ? todo : Utils.extend({}, todo, { title: text });
        });

        this.inform();
    }

    public clearCompleted() {
        this.todos = this.todos.filter(todo => {
            if (todo.completed) {
                this._outlookTasks.delete(todo);
            }
            return !todo.completed;
        });

        this.inform();
    }
}