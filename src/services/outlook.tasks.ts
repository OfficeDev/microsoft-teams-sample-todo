import { ITodo, IProfile } from '../core';
import { Authenticator, Storage, IToken, IEndpoint } from '../auth';

interface ITodoService {
    authenticator: Authenticator;
    login: () => Promise<IToken>;
    get: () => Promise<ITodo[]>;
    create: (todo: ITodo) => Promise<ITodo>;
    delete: (todo: ITodo) => Promise<boolean>;
    markAsComplete: (todo: ITodo) => Promise<ITodo>;
}

export class OutlookTasks implements ITodoService {
    graphToken: IToken;
    tasksToken: IToken;
    private _retryCounter = 0;
    private _baseUrl: string = 'https://outlook.office.com/api/v2.0';
    private _graphUrl: string = 'https://graph.microsoft.com/v1.0';

    authenticator: Authenticator;
    storage: Storage<IProfile>;

    constructor(scope?: string) {
        this.authenticator = new Authenticator();
        this.storage = new Storage<IProfile>('GraphProfile');
        this.authenticator.endpoints.registerMicrosoftAuth('73d044ea-4ae0-4bd8-b26c-7c2f924410a2', {
            redirectUrl: 'https://localhost:3000/config.html'
        });

        this.authenticator.endpoints.add('Tasks', {
            clientId: '73d044ea-4ae0-4bd8-b26c-7c2f924410a2',
            baseUrl: 'https://login.microsoftonline.com/common/oauth2/v2.0',
            redirectUrl: 'https://localhost:3000/config.html',
            authorizeUrl: '/authorize',
            responseType: 'token',
            scope: 'https://outlook.office.com/tasks.readwrite',
            nonce: true,
            state: true,
            extraParameters: '&response_mode=fragment',
        } as IEndpoint)

        this.graphToken = this.authenticator.tokens.get('Microsoft');
    }

    login(graph?: boolean): Promise<IToken> {
        var scope = graph ? 'Microsoft' : 'Tasks';
        return this.authenticator.authenticate('Tasks')
            .then(token => this.tasksToken = token)
            .catch(error => {
                console.error(error);
                throw new Error('Failed to login using your Microsoft Account');
            });
    }

    logout() {
        this.authenticator.tokens.clear();
        this.storage.clear();
    }

    profile(): Promise<IProfile> {
        var profile = this.storage.get('Profile');
        if (profile) return Promise.resolve(profile);
        return this._isAuthenticated()
            .then(token => this._getProfile())
            .then(profile => this.storage.add('Profile', profile));
    }

    private _getProfile(): Promise<IProfile> {
        var profile = {} as IProfile;
        return Promise.all([
            this._getUserInfo(),
            this._getUserPhoto()
        ])
            .then(response => {
                var [userInfo, userPhoto] = response;
                profile = userInfo;
                profile.thumbnail = userPhoto;
                return profile;
            });
    }

    private _getUserInfo(): Promise<any> {
        return fetch(`${this._graphUrl}/me`,
            {
                headers: {
                    Authorization: `Bearer ${this.graphToken.access_token}`
                }
            })
            .then(res => res.json());
    }

    private _getUserPhoto(): Promise<any> {
        return fetch(`${this._graphUrl}/me/photo/$value`,
            {
                headers: {
                    Authorization: `Bearer ${this.graphToken.access_token}`
                }
            })
            .then(res => res.blob())
            .then(blob => window.URL.createObjectURL(blob))
    }

    get(): Promise<ITodo[]> {
        return this._isAuthenticated().then(token => this._getTodos());
    }

    private _getTodos(): Promise<ITodo[]> {
        return fetch(`${this._baseUrl}/me/tasks`,
            {
                headers: {
                    Authorization: `Bearer ${this.tasksToken.access_token}`
                }
            })
            .then(res => res.json())
            .then(res => res.value.map(this._convertTask))
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
                    'Authorization': `Bearer ${this.tasksToken.access_token}`,
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            })
            .then(res => res.json())
            .then(this._convertTask)
            .catch(error => {
                console.error(error);
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
                    'Authorization': `Bearer ${this.tasksToken.access_token}`,
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            })
            .then(res => res.json())
            .then(this._convertTask)
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
                    'Authorization': `Bearer ${this.tasksToken.access_token}`,
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
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
                    'Authorization': `Bearer ${this.tasksToken.access_token}`,
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            })
            .then(res => res.json())
            .then(this._convertTask)
            .catch(error => {
                console.error(error);
                throw new Error('Failed to mark your todo as complete');
            });
    }

    private _isAuthenticated(graph?: boolean): Promise<IToken> {
        var token = graph ? this.graphToken : this.tasksToken;
        if (token == null) {
            if (this._retryCounter < 3) {
                return this.login()
                    .catch(e => {
                        console.error(`Login failed... retrying... ${++this._retryCounter}`);
                        return this._isAuthenticated(graph);
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