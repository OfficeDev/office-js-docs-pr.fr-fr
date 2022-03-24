---
title: Interface utilisateur Fluent - Comment faire pour les modules add-in Office ?
description: Découvrez comment utiliser les Fluent’interface utilisateur React dans Office de l’interface utilisateur.
ms.date: 01/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 453befe44dbcec6527930fcd73c5cb2cb243d965
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743131"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a>Utiliser Fluent’interface utilisateur React dans les Office de l’interface utilisateur

Fluent interface utilisateur React est l’infrastructure frontale JavaScript open source officielle conçue pour créer des expériences qui s’intègrent parfaitement à un large éventail de produits Microsoft, y compris Office. Il fournit des composants robustes, à jour, accessibles et basés sur React, qui sont hautement personnalisables à l'aide de CSS-in-JS.

> [!NOTE]
> Cet article décrit l’utilisation des Fluent’interface utilisateur React dans le contexte de Office de l’interface utilisateur. Mais il est également utilisé dans un large éventail d’applications Microsoft 365 et d’extensions. Pour plus d’informations, [voir Fluent’interface utilisateur React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) et le repo open source [Fluent UI Web](https://github.com/microsoft/fluentui).

Cet article explique comment créer un module qui est créé avec React et qui utilise Fluent’interface utilisateur React composants.

## <a name="create-an-add-in-project"></a>Création d’un projet de complément

Vous utiliserez le générateur Yeoman pour les compléments Office pour créer un projet de complément utilisant React.

### <a name="install-the-prerequisites"></a>Installez les composants requis

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a>Créez le projet

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Sélectionnez un type de projet :** `Office Add-in Task Pane project using React framework`
- **Sélectionnez un type de script :** `TypeScript`
- **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
- **Quelle application client Office voulez-vous prendre en charge ?** `Word`

![Capture d’écran montrant les invites et réponses relatives au générateur Yeoman dans une interface de ligne de commande.](../images/yo-office-word-react.png)

Après avoir exécuté l’Assistant, le générateur crée le projet et installe les composants Node de prise en charge.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a>Essayez

1. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > Si vous testez votre complément sur Mac, exécutez la commande suivante avant de continuer. Lorsque vous exécutez cette commande, le serveur web local démarre.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - Pour tester votre complément dans Word, exécutez la commande suivante dans le répertoire racine de votre projet. Cela démarre le serveur web local et ouvre Word avec votre add-in chargé.

        ```command&nbsp;line
        npm start
        ```

    - Pour tester votre complément dans Word sur un navigateur, exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre. Remplacez « {url} » par l’URL d’un document Word sur votre OneDrive ou une bibliothèque SharePoint sur laquelle vous avez des autorisations.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

3. Pour ouvrir le volet Des tâches du add-in, sous **l’onglet Accueil** , sélectionnez le bouton Afficher **le volet Des** tâches. Remarquez le texte par défaut et le bouton **Exécuter** en bas du volet Office. Dans le reste de cette walkthrough, vous redéfinirez ce texte et ce bouton en créant un composant React qui utilise des composants UX à partir de Fluent’interface utilisateur React.

    ![Screenshot showing the Word application with the Show Taskpane ribbon button highlighted and the Run button and immediately preceding text highlighted in the task pane.](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a>Créer un composant React qui utilise Fluent’interface utilisateur React

À ce stade, vous avez créé un complément très rudimentaire du volet Office standard en utilisant React. Ensuite, procédez comme suit pour créer un nouveau composant React (`ButtonPrimaryExample`) dans le projet de complément. Le composant utilise les composants `Label` de `PrimaryButton` la Fluent’interface utilisateur React.

1. Ouvrez le dossier du projet créé par le générateur Yeoman et accédez à **src\taskpane\components**.
2. Dans ce dossier, créez un fichier nommé **Button.tsx**.
3. Dans **Button.tsx**, ajoutez le code suivant pour définir le composant `ButtonPrimaryExample`.

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from '@fluentui/react/lib/components/Button';
import { Label } from '@fluentui/react/lib/components/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor(props) {
    super(props);
  }

  insertText = async () => {
    // In the click event, write text to the document.
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph('Hello Fluent UI React!', Word.InsertLocation.end);
      await context.sync();
    });
  }

  public render() {
    let { disabled } = this.props;
    return (
      <div className='ms-BasicButtonsExample'>
        <Label>Click the button to insert text.</Label>
        <PrimaryButton
          data-automation-id='test'
          disabled={ disabled }
          text='Insert text...'
          onClick={ this.insertText } />
      </div>
    );
  }
}
```

Ce code effectue les opérations suivantes :

- Fait référence à la bibliothèque React en utilisant `import * as React from 'react';`.
- Fait référence à Fluent’interface utilisateur React composants (`PrimaryButton`, `IButtonProps`, `Label`) qui sont utilisés pour créer`ButtonPrimaryExample`.
- Déclare le nouveau composant `ButtonPrimaryExample` en utilisant `export class ButtonPrimaryExample extends React.Component`.
- Déclare la fonction `insertText` qui gère l’événement du bouton `onClick`.
- Définit l’interface utilisateur du composant React dans la fonction `render`. Le code HTML utilise `Label` `PrimaryButton` les composants de Fluent interface `onClick` utilisateur React et spécifie que lorsque l’événement se déclenche, `insertText` la fonction s’exécute.

## <a name="add-the-react-component-to-your-add-in"></a>Ajoutez le composant React à votre complément

Ajoutez le `ButtonPrimaryExample` composant à votre application en ouvrant **src\components\App.tsx** et en effectuant les étapes suivantes.

1. Ajoutez l’instruction importation suivante pour référencer `ButtonPrimaryExample` dans **Button.tsx**.

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. Supprimez l’instruction import suivante.

    ```typescript
    import Progress from './Progress';
    ```

3. Remplacez la fonction `render()` par défaut par le code suivant qui utilise `ButtonPrimaryExample`.

    ```typescript
    render() {
      return (
        <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what this add-in can do for you today!" items={this.state.listItems} >
          <ButtonPrimaryExample />
        </HeroList>
        </div>
      );
    }
    ```

4. Enregistrez les modifications apportées à **App.tsx**.

## <a name="see-the-result"></a>Regardez le résultat

Dans Word, le volet Office complément se met automatiquement à jour lorsque vous enregistrez les modifications apportées à **App.tsx**. Le texte et le bouton par défaut en bas du volet Office indiquent désormais l’interface utilisateur définie par le composant `ButtonPrimaryExample`. Sélectionnez le bouton **Insérer un texte...** pour insérer du texte dans le document.

![Capture d’écran montrant l’application Word avec « Insérer du texte... » bouton et texte qui précède immédiatement mis en surbrill](../images/word-task-pane-with-react-component.png)

Félicitations, vous avez créé un add-in de volet de tâches à l’aide de React et Fluent’interface React !

## <a name="see-also"></a>Voir aussi

- [Word Add-in GettingStartedFabricReact](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [Cœur de fabric dans les modules](fabric-core.md)
- [Modèles de conception de l’expérience utilisateur pour les compléments Office](ux-design-pattern-templates.md)
