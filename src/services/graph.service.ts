import { ITodo } from '../core';
import { Authenticator, IToken } from '@microsoft/office-js-helpers';

interface ITodoService {
    login: () => Promise<IToken>;
    get: () => Promise<ITodo[]>;
    create: (todo: ITodo) => Promise<ITodo>;
    delete: (todo: ITodo) => Promise<boolean>;
    markAsComplete: (todo: ITodo) => Promise<ITodo>;
}

export class GraphService implements ITodoService {
    private _token: IToken;
    private _retryCounter = 0;
    private _baseUrl: string = 'https://outlook.office.com/api/v2.0';

    constructor(public authenticator: Authenticator) {
        authenticator.endpoints.registerMicrosoftAuth('73d044ea-4ae0-4bd8-b26c-7c2f924410a2', {
            scope: 'https://outlook.office.com/tasks.readwrite'
        });
    }

    login(): Promise<IToken> {
        return this.authenticator.authenticate('microsoft')
            .then(token => this._token = token)
            .catch(error => {
                console.error(error);
                throw new Error('Failed to login using your Microsoft Account');
            });
    }

    get(): Promise<ITodo[]> {
        return this._isAuthenticated().then(token => this._getTodos());
    }

    private _getTodos(): Promise<ITodo[]> {
        return fetch(`${this._baseUrl}/me/tasks`,
            {
                headers: {
                    Authorization: `Bearer ${this._token.access_token}`
                }
            })
            .then(res => res.json())
            .then(res => res.value.map(this._convertToDo))
            .catch(error => {
                console.error(error);
                throw new Error('Failed to get your todos');
            });
    }

    create(todo: ITodo): Promise<ITodo> {
        return this._isAuthenticated().then(token => this._createTodo(todo));
    }

    private _createTodo(todo: ITodo): Promise<ITodo> {
        return fetch(`${this._baseUrl}/me/tasks`,
            {
                method: "POST",
                body: JSON.stringify({
                    Subject: todo.title,
                    Status: todo.completed ? 'Completed' : 'NotStarted',
                    Importance: todo.importance == null ? 'normal' : todo.importance
                }),
                headers: {
                    Authorization: `Bearer ${this._token.access_token}`
                }
            })
            .then(res => res.json())
            .then(this._convertToDo)
            .catch(error => {
                console.error(error);
                throw new Error('Failed to create your todo');
            });
    }

    delete(todo: ITodo): Promise<boolean> {
        return this._isAuthenticated().then(token => this._deleteTodo(todo));
    }

    private _deleteTodo(todo: ITodo): Promise<boolean> {
        return fetch(`${this._baseUrl}/me/tasks('${todo.id}')`,
            {
                method: "DELETE",
                headers: {
                    Authorization: `Bearer ${this._token.access_token}`
                }
            })
            .then(res => res.status === 204)
            .catch(error => {
                console.error(error);
                throw new Error('Failed to delete your todo');
            });
    }

    markAsComplete(todo: ITodo): Promise<ITodo> {
        return this._isAuthenticated().then(token => this._markAsComplete(todo));
    }

    private _markAsComplete(todo: ITodo): Promise<ITodo> {
        return fetch(`${this._baseUrl}/me/tasks('${todo.id}')/complete`,
            {
                method: "POST",
                headers: {
                    Authorization: `Bearer ${this._token.access_token}`
                }
            })
            .then(res => res.json())
            .then(this._convertToDo)
            .catch(error => {
                console.error(error);
                throw new Error('Failed to mark your todo as complete');
            });
    }

    private _isAuthenticated(): Promise<IToken> {
        if (this._token == null) {
            if (this._retryCounter < 3) {
                return this.login()
                    .catch(e => {
                        console.error(`Login failed... retrying... ${++this._retryCounter}`);
                        return this._isAuthenticated();
                    });
            }
            else {
                this._retryCounter = 0;
                throw new Error('Failed to login using your Microsoft Account');
            }
        }
        else {
            return Promise.resolve(this._token);
        }
    }

    private _convertToDo(task: any): ITodo {
        return {
            id: task.Id,
            title: task.Subject,
            completed: task.Status === 'Completed',
            importance: task.Importance.toLowerCase()
        } as ITodo;
    }
}