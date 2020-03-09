---
page_type: sample
products:
- office-teams
- office-outlook
- office-365
languages:
- typescript
- nodejs
extensions:
  contentType: samples
  technologies:
  - Tabs
  createdDate: 10/19/2016 12:06:51 PM
---
# Exemplo de aplicativo de guias “Lista de tarefas” do Microsoft Teams
[![Status da Compilação](https://travis-ci.org/OfficeDev/microsoft-teams-sample-todo.svg?branch=master)](https://travis-ci.org/OfficeDev/microsoft-teams-sample-todo)

Este é um exemplo de [aplicativo de guias do Microsoft Teams](https://aka.ms/microsoftteamstabsplatform). O intuito do exemplo é ilustrar como é simples converter um aplicativo Web em um aplicativo de guias do Microsoft Teams. O aplicativo Web existente, [**TodoMVC for React**](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react) fornece um gerenciador de tarefas básico que se integra às suas Tarefas do Outlook pessoais. Com apenas algumas pequenas modificações, essa exibição da Web pode ser adicionada a um canal como um aplicativo de guias. Dê uma olhada na [diferença de código entre as ramificações “antes” e “depois”](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after) para ver as alterações feitas.

> **Observação:** Este não é um exemplo realista de um aplicativo de colaboração em equipe. As tarefas mostradas pertencem à conta individual do usuário e não a uma conta compartilhada de equipe.

**Para saber mais sobre o desenvolvimento de experiências para o Microsoft Teams, confira a [documentação para desenvolvedores](https://msdn.microsoft.com/en-us/microsoft-teams/index) do Microsoft Teams.**

## Pré-requisitos

1. Uma [conta do Office 365 com acesso ao Microsoft Teams](https://msdn.microsoft.com/en-us/microsoft-teams/setup).
2. Este exemplo é criado usando o [Node.js](https://nodejs.org). Baixe e instale a versão recomendada, se ainda não a tiver.

## Executar o aplicativo

### Hospedar os aplicativos de guias “página de configuração” e “página de conteúdo”

Como a configuração de guias e páginas de conteúdo são aplicativos Web, eles precisam estar hospedados.  

Primeiro, vamos preparar o projeto:

1. Clone o repositório.
2. Abra uma linha de comando no subdiretório do repositório.
3. Execute o `npm install` (incluído como parte do Node.js) na linha de comandos.

Para hospedar o aplicativo localmente:
* No subdiretório do repositório, execute `npm start` para iniciar o `webpack-dev-server` e permitir o funcionamento do aplicativo para desenvolvedores.
* Observe que o manifesto do desenvolvedor está configurado para usar o localhost, portanto, nenhuma configuração adicional seria necessária.

Opcionalmente, para criar um aplicativo implantável, que você pode hospedar em seu próprio ambiente:
* No subdiretório do repositório, execute `npm run build` para gerar uma compilação implantável.
* Modifique o config.url no manifesto do produto para direcionar para o seu local de hospedagem. [Confira a seguir](#registering-an-application-to-authenticate-with-microsoft)

### Adicionar a guia ao Microsoft Teams

1. Baixe o arquivo zip do [pacote para desenvolvedores de aplicativos de guias](https://github.com/OfficeDev/microsoft-teams-sample-todo/tree/master/package/todo.dev.zip) para esta amostra.
2. Crie uma nova equipe para teste, se necessário. Clique em **Criar equipe** na parte inferior do painel à esquerda.
3. Selecione a equipe no painel à esquerda, selecione **... (mais opções)** e, em seguida, selecione **Exibir equipe**.
4. Selecione a guia **Desenvolvedor (Visualização)** e, em seguida, selecione **Carregar**.
5. Navegue até o arquivo zip baixado na etapa 1 acima e selecione-o.
6. Vá para qualquer canal da equipe. Clique no sinal “+” à direita das guias existentes.
7. Selecione sua guia na galeria exibida.
8. Aceite a solicitação de consentimento.
9. Se necessário, entre usando sua conta corporativa/de estudante do Office 365. Observe que o código tentará fazer autenticação silenciosa, se possível.
10. Valide as informações de autenticação.
11. Pressione Salvar para adicionar a guia ao canal.

> **Observação:** Para carregar novamente um pacote atualizado com a mesma `id`, clique no ícone “Substituir” no final da linha da tabela da guia. Não clique em “Carregar” novamente: o Microsoft Teams irá informar que a guia já existe.

> É recomendável ter várias configurações, uma por ambiente. Os nomes dos arquivos zip podem conter qualquer coisa como `todo.dev.zip`, `todo.prod.zip` etc., mas o zip precisa conter um `manifest.json` com uma `id` exclusiva. Confira [Criar um manifesto para a sua guia](https://msdn.microsoft.com/en-us/microsoft-teams/createpackage).

## Caminho para o código

Enquanto a ramificação `mestre` mostra o estado mais recente da amostra, observe a [diferença de código](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after) a seguir entre:

* [`antes`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/before): o aplicativo inicial

* [`depois`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/after): o aplicativo após a integração com o Microsoft Teams.

Etapa por etapa:

1. Adicionamos uma nova página `config.html` e `config.tsx`, responsável pelo aplicativo, para permitir que o usuário manipule todas as configurações, realize a autenticação de logon único, etc. durante a primeira inicialização. Isso é necessário para que o administrador da equipe possa definir o aplicativo e as configurações. Confira [Criar a página de configuração](https://msdn.microsoft.com/en-us/microsoft-teams/createconfigpage).

2. Adicionamos o mesmo arquivo `config.html` à nossa configuração `webpack.common.js` para que ele possa injetar os pacotes corretos durante o tempo de execução.

3. Adicionamos uma referência ao `MicrosoftTeams.js` em nosso index.html e adicionamos `MicrosoftTeams.d.ts` para o Typescript intellisense.

4. Adicionamos `manifest.dev.json`, `manifest.prod.json` e dois logotipos para os tamanhos *44x44* e *88x88*. Lembre-se de renomeá-los como `manifest.json` nos arquivos zip que você carrega no Microsoft Teams.

5. Adicionamos alguns estilos ao nosso `app.css`.

6. Finalmente, modificamos a `autenticação` em outlook.tasks.ts para depender, em vez disso, de “useMicrosoftTeamsAuth”, um novo recurso da versão beta do OfficeHelpers mencionada neste exemplo.

### Invocar a caixa de diálogo Autenticação

Quando o usuário adiciona a guia, a página de configuração é apresentada (config.html). Nesse caso, o código autentica o usuário, se possível.

A autenticação utiliza uma nova função específica do Teams na versão mais recente (> 0.4.0) da biblioteca [office-js-helpers](https://github.com/OfficeDev/office-js-helpers). Essa função auxiliar tentará fazer uma autenticação silenciosa, mas, se não conseguir, invocará para você a caixa de diálogo de autenticação específica do Microsoft Teams. Para saber mais sobre o processo de autenticação completo no Microsoft Teams, confira [Autenticar um usuário](https://msdn.microsoft.com/en-us/microsoft-teams/auth) na [documentação para desenvolvedores](https://msdn.microsoft.com/en-us/microsoft-teams/index) do Microsoft Teams.

Em outlook.tasks.ts:
```typescript
 return this.authenticator.authenticate('Microsoft')
    .then(token => this._token = token)
    .catch(error => {
        Utilities.log(error);
        throw new Error('Failed to login using your Microsoft Account');
    });
```

### Manipular o evento “Salvar”

Após a entrada bem-sucedida, o usuário salvará a guia no canal. O código a seguir ativa o botão Salvar e define o SaveHandler, que armazenará o conteúdo a ser exibido na guia (nesse caso, apenas o index.html do projeto).

Em config.tsx:
```typescript
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
```

## Tecnologia usada

Esse exemplo usa o seguinte:

1. [`React by Facebook`](https://facebook.github.io/react/) como a UI Framework.
2. [`TypeScript`](https://www.typescriptlang.org/) como o transcompilador.
4. Base [`TodoMVC`](http://todomvc.com/examples/typescript-react/#/) para a funcionalidade do TodoMVC.
5. [`Webpack`](https://webpack.github.io/) como a ferramenta de compilação.

## Recursos adicionais

### Registrar um aplicativo para autenticar com o Microsoft
Observe que isso não será necessário se você usar a opção Desenvolvedor local acima, mas se optar por hospedar essa guia em seu próprio ambiente, deverá registrar o aplicativo para fazer a autenticação.

1. Vá até o [Portal de Registro de Aplicativos da Microsoft](https://apps.dev.microsoft.com).
2. Entre com sua conta corporativa/de estudante do Office 365. Não use sua conta pessoal da Microsoft.
2. Adicione um novo aplicativo.
2. Anote sua nova `ID do aplicativo`.
2. Clique em `Adicionar plataforma` e escolha `Web`.
3. Selecione `Permitir fluxo implícito` e configure a URL de redirecionamento para `https://<mywebsite>/config.html`.

Para saber mais sobre como hospedar suas próprias guias, confira o [LEIAME “Introdução” do Microsoft Teams como exemplo](https://github.com/OfficeDev/microsoft-teams-sample-get-started#host-tab-pages-over-https).

>**Observação:** Por padrão, sua organização deve permitir que você crie novos aplicativos. Caso contrário, você pode obter um locatário de teste de desenvolvedor, uma assinatura de avaliação de um ano do Desenvolvedor do Office 365 sem nenhum custo. [Confira como](https://msdn.microsoft.com/en-us/microsoft-teams/setup).


## Créditos

Este projeto é baseado no modelo TodoMVC Typescript - React localizado [aqui](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react).

## Colaboração

Leia [Colaboração](contributing.md) para obter detalhes sobre o nosso código de conduta e o processo de envio de solicitações para nós.

Este projeto adotou o [Código de Conduta de Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/).  Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.

## Controle de Versão

Usamos o [SemVer](http://semver.org/) para controle de versão. Para verificar as versões disponíveis, confira as [tags neste repositório](https://github.com/officedev/microsoft-teams-sample-todo/tags).

## Licença

Este projeto está licenciado sob a Licença MIT, confira o arquivo [Licença](LICENSE) para obter detalhes
