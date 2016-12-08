import { ITodo, IProfile } from '../core';

/**
 * We are using OfficeHelpers library as it allows us to complete authentication
 * with relative ease and also provides other useful Utilities.
 *
 * Note: We have included a beta version of OfficeHelpers that has support for
 * MicrosoftTeams and the API signatures might change when OfficeHelpers for
 * Microsoft Teams releases.
 */
import { Authenticator, IToken, Utilities } from '@microsoft/office-js-helpers';

interface ITodoService {
    authenticator: Authenticator;
    login: () => Promise<IToken>;
    get: () => Promise<ITodo[]>;
    create: (todo: ITodo) => Promise<ITodo>;
    delete: (todo: ITodo) => Promise<boolean>;
    markAsComplete: (todo: ITodo) => Promise<ITodo>;
}

export class OutlookTasks implements ITodoService {
    private _token: IToken;
    private _retryCounter = 0;
    private _baseUrl: string = 'https://outlook.office.com/api/v2.0';
    private _login: boolean;

    authenticator: Authenticator;

    /**
     * Note: This is a demo clientID and can be removed/deactivated
     * at any point of time. Please modify this to use your own
     * clientId.
     */
    clientId = 'bd9464a2-8b29-4165-a3c8-4d4dfd64b59a';

    constructor() {

        /**
         * OfficeHelpers provides a authentication utility that does Implicit OAuth
         * in simple steps and supports Microsoft, AzureAD, Google, Facebook out of the box.
         *
         * Note: We have included a beta version of OfficeHelpers that has support for
         * MicrosoftTeams and the API signatures might change when OfficeHelpers for
         * Microsoft Teams releases.
         *
         * If you wish to achieve the same authentication call without OfficeHelpers,
         * then you would need to call the following Microsoft Teams API manually and
         * handle token mangement yourself.
         *
         * microsoftTeams.authentication.authenticate({
         *     url: <constructed OAuth url to the authentication provider with required parameters>,
         *     width: <width in pixels>,
         *     height: <height in pixels>,
         *     successCallback: <handle message from dialog here>,
         *     failureCallback: <handle error from dialog here>
         * });
         *
         * Inside the dialog you can call
         * microsoftTeams.authentication.notifySuccess(<message string to be sent back to app>)
         * OR
         * microsoftTeams.authentication.notifyFailure(<error string to be sent back to app>)
         * accordingly.
         */
        this.authenticator = new Authenticator();
        this.authenticator.endpoints.registerMicrosoftAuth(this.clientId, {
            baseUrl: 'https://login.microsoftonline.com/organizations/oauth2/v2.0',
            scope: 'https://outlook.office.com/tasks.readwrite'
        });


        this._token = this.authenticator.tokens.get('Microsoft');
        if (this._token == null) {
            this.login();
        }
    }

    login(): Promise<IToken> {
        /**
         * Read the comments above to achieve Authentication without depending on
         * the OfficeHelpers library.
         */
        return this.authenticator.useMicrosoftTeamsAuth('Microsoft')
            .then(token => this._token = token)
            .catch(error => {
                Utilities.log(error);
                throw new Error('Failed to login using your Microsoft Account');
            });
    }

    logout() {
        this.authenticator.tokens.clear();
    }

    get(): Promise<ITodo[]> {
        return this._isAuthenticated().then(token => this._getTodos());
    }

    private _checkForErrors = response => {
        if (response.error) {
            throw new Error(response.error.message)
        }

        return response;
    }

    private _getTodos(): Promise<ITodo[]> {
        return fetch(`${this._baseUrl}/me/tasks`,
            {
                headers: {
                    'Authorization': `Bearer ${this._token.access_token}`
                }
            })
            .then(res => res.json())
            .then(res => this._checkForErrors(res))
            .then(res => res.value.map(this._convertTask))
            .catch(error => {
                Utilities.log(error);
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
                    'Authorization': `Bearer ${this._token.access_token}`,
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            })
            .then(res => res.json())
            .then(res => this._checkForErrors(res))
            .then(this._convertTask)
            .catch(error => {
                Utilities.log(error);
                throw new Error('Failed to create your todo');
            });
    }

    update(todo: ITodo): Promise<ITodo> {
        return this._isAuthenticated().then(token => this._updateTodo(todo));
    }

    private _updateTodo(todo: ITodo): Promise<ITodo> {
        return fetch(`${this._baseUrl}/me/tasks('${todo.id}')`,
            {
                method: "PATCH",
                body: JSON.stringify({
                    Subject: todo.title,
                    Status: todo.completed ? 'Completed' : 'NotStarted',
                    Importance: todo.importance == null ? 'normal' : todo.importance
                }),
                headers: {
                    'Authorization': `Bearer ${this._token.access_token}`,
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            })
            .then(res => res.json())
            .then(res => this._checkForErrors(res))
            .then(this._convertTask)
            .catch(error => {
                Utilities.log(error);
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
                    'Authorization': `Bearer ${this._token.access_token}`,
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            })
            .then(res => res.status === 204)
            .then(res => this._checkForErrors(res))
            .catch(error => {
                Utilities.log(error);
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
                    'Authorization': `Bearer ${this._token.access_token}`,
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            })
            .then(res => res.json())
            .then(res => this._checkForErrors(res))
            .then(this._convertTask)
            .catch(error => {
                Utilities.log(error);
                throw new Error('Failed to mark your todo as complete');
            });
    }

    private _isAuthenticated(): Promise<IToken> {
        var token = this._token;
        if (token == null) {
            if (this._retryCounter < 3) {
                return this.login()
                    .catch(e => {
                        Utilities.log(`Login failed... retrying... ${++this._retryCounter}`);
                        return this._isAuthenticated();
                    });
            }
            else {
                this._retryCounter = 0;
                throw new Error('Failed to login using your Microsoft Account');
            }
        }
        else {
            return Promise.resolve(token);
        }
    }

    private _convertTask(task: any): ITodo {
        return {
            id: task.Id,
            title: task.Subject,
            completed: task.Status === 'Completed',
            importance: task.Importance && task.Importance.toLowerCase()
        } as ITodo;
    }
}