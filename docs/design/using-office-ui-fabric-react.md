---
title: Interface utilisateur Fluent React dans Office de l’interface utilisateur
description: Découvrez comment utiliser l’interface utilisateur Fluent React dans Office de l’interface utilisateur.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: cb7f04c21a52a2e4a3f271abc56aa325dd2b02fd
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330141"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a><span data-ttu-id="33139-103">Utiliser l’interface utilisateur Fluent React dans Office de l’interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="33139-103">Use Fluent UI React in Office Add-ins</span></span>

<span data-ttu-id="33139-104">Fluent UI React est l’infrastructure frontale JavaScript open source officielle conçue pour créer des expériences qui s’intègrent parfaitement à un large éventail de produits Microsoft, notamment Office.</span><span class="sxs-lookup"><span data-stu-id="33139-104">Fluent UI React is the official open-source JavaScript front-end framework designed to build experiences that fit seamlessly into a broad range of Microsoft products, including Office.</span></span> <span data-ttu-id="33139-105">Il fournit des composants robustes, à jour et accessibles React qui sont hautement personnalisables à l’aide de CSS-in-JS.</span><span class="sxs-lookup"><span data-stu-id="33139-105">It provides robust, up-to-date, accessible React-based components which are highly customizable using CSS-in-JS.</span></span>

> [!NOTE]
> <span data-ttu-id="33139-106">Cet article décrit l’utilisation de l’interface utilisateur Fluent React dans le contexte de Office de l’interface utilisateur. Mais il est également utilisé dans un large éventail d’applications Microsoft 365 et d’extensions.</span><span class="sxs-lookup"><span data-stu-id="33139-106">This article describes the use of Fluent UI React in the context of Office Add-ins. But it is also used in a wide range of Microsoft 365 apps and extensions.</span></span> <span data-ttu-id="33139-107">Pour plus d’informations, consultez [la React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) interface utilisateur Fluent et le site web d’interface utilisateur [Fluent open](https://github.com/microsoft/fluentui)source.</span><span class="sxs-lookup"><span data-stu-id="33139-107">For more information, see [Fluent UI React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) and the open source repo [Fluent UI Web](https://github.com/microsoft/fluentui).</span></span>

<span data-ttu-id="33139-108">Cet article explique comment créer un add-in qui est créé à l’React et qui utilise l’interface utilisateur Fluent React composants.</span><span class="sxs-lookup"><span data-stu-id="33139-108">This article describes how to create an add-in that's built with React and uses Fluent UI React components.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="33139-109">Création d’un projet de complément</span><span class="sxs-lookup"><span data-stu-id="33139-109">Create an add-in project</span></span>

<span data-ttu-id="33139-110">Vous utiliserez le générateur Yeoman pour les compléments Office pour créer un projet de complément utilisant React.</span><span class="sxs-lookup"><span data-stu-id="33139-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="33139-111">Installez les composants requis</span><span class="sxs-lookup"><span data-stu-id="33139-111">Install the prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="33139-112">Créez le projet</span><span class="sxs-lookup"><span data-stu-id="33139-112">Create the project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="33139-113">**Sélectionnez un type de projet :** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="33139-113">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="33139-114">**Sélectionnez un type de script :** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="33139-114">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="33139-115">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="33139-115">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="33139-116">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="33139-116">**Which Office client application would you like to support?**</span></span> `Word`

![Capture d’écran montrant les invites et réponses relatives au générateur Yeoman dans une interface de ligne de commande](../images/yo-office-word-react.png)

<span data-ttu-id="33139-118">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="33139-118">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a><span data-ttu-id="33139-119">Essayez</span><span class="sxs-lookup"><span data-stu-id="33139-119">Try it out</span></span>

1. <span data-ttu-id="33139-120">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="33139-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="33139-121">Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="33139-121">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="33139-122">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="33139-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="33139-123">Si vous êtes invité à installer un certificat après avoir exécuté une des commandes suivantes, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="33139-123">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="33139-124">Il se peut également que vous deviez exécuter votre invite de commande ou votre terminal en tant qu'administrateur pour que les modifications soient effectuées.</span><span class="sxs-lookup"><span data-stu-id="33139-124">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    > [!TIP]
    > <span data-ttu-id="33139-125">Si vous testez votre complément sur Mac, exécutez la commande suivante avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="33139-125">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="33139-126">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="33139-126">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="33139-127">Pour tester votre complément dans Word, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="33139-127">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="33139-128">Cela a pour effet de démarrer le serveur web local (s’il n’est pas déjà en cours d’exécution) et d’ouvrir Word avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="33139-128">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="33139-129">Pour tester votre complément dans Word sur un navigateur, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="33139-129">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="33139-130">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="33139-130">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="33139-131">Pour utiliser votre complément, ouvrez un nouveau document dans Word sur le web, puis chargez la version test de votre complément en suivant les instructions de l’article relatif au [chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="33139-131">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="33139-132">Pour ouvrir le volet Des tâches  du add-in, sous l’onglet Accueil, sélectionnez le bouton Afficher le **volet Des** tâches.</span><span class="sxs-lookup"><span data-stu-id="33139-132">To open the add-in task pane, on the **Home** tab, choose the **Show Taskpane** button.</span></span> <span data-ttu-id="33139-133">Remarquez le texte par défaut et le bouton **Exécuter** en bas du volet Office.</span><span class="sxs-lookup"><span data-stu-id="33139-133">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="33139-134">Dans le reste de cette walkthrough, vous redéfinirez ce texte et ce bouton en créant un composant React qui utilise des composants UX à partir de fluent UI React.</span><span class="sxs-lookup"><span data-stu-id="33139-134">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fluent UI React.</span></span>

    ![Screenshot showing the Word application with the Show Taskpane ribbon button highlighted and the Run button and immediately preceding text highlighted in the task pane](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a><span data-ttu-id="33139-136">Créer un composant React qui utilise l’interface utilisateur Fluent React</span><span class="sxs-lookup"><span data-stu-id="33139-136">Create a React component that uses Fluent UI React</span></span>

<span data-ttu-id="33139-137">À ce stade, vous avez créé un complément très rudimentaire du volet Office standard en utilisant React.</span><span class="sxs-lookup"><span data-stu-id="33139-137">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="33139-138">Ensuite, procédez comme suit pour créer un nouveau composant React (`ButtonPrimaryExample`) dans le projet de complément.</span><span class="sxs-lookup"><span data-stu-id="33139-138">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="33139-139">Le composant utilise les composants de l’interface utilisateur `Label` `PrimaryButton` Fluent React.</span><span class="sxs-lookup"><span data-stu-id="33139-139">The component uses the `Label` and `PrimaryButton` components from Fluent UI React.</span></span>

1. <span data-ttu-id="33139-140">Ouvrez le dossier du projet créé par le générateur Yeoman et accédez à **src\taskpane\components**.</span><span class="sxs-lookup"><span data-stu-id="33139-140">Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.</span></span>
2. <span data-ttu-id="33139-141">Dans ce dossier, créez un fichier nommé **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="33139-141">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="33139-142">Dans **Button.tsx**, ajoutez le code suivant pour définir le composant `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="33139-142">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

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

<span data-ttu-id="33139-143">Ce code effectue les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="33139-143">This code does the following:</span></span>

- <span data-ttu-id="33139-144">Fait référence à la bibliothèque React en utilisant `import * as React from 'react';`.</span><span class="sxs-lookup"><span data-stu-id="33139-144">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="33139-145">Fait référence à l’interface utilisateur Fluent React composants ( `PrimaryButton` , , ) qui sont utilisés pour créer `IButtonProps` `Label` `ButtonPrimaryExample` .</span><span class="sxs-lookup"><span data-stu-id="33139-145">References the Fluent UI React components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="33139-146">Déclare le nouveau composant `ButtonPrimaryExample` en utilisant `export class ButtonPrimaryExample extends React.Component`.</span><span class="sxs-lookup"><span data-stu-id="33139-146">Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="33139-147">Déclare la fonction `insertText` qui gère l’événement du bouton `onClick`.</span><span class="sxs-lookup"><span data-stu-id="33139-147">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="33139-148">Définit l’interface utilisateur du composant React dans la fonction `render`.</span><span class="sxs-lookup"><span data-stu-id="33139-148">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="33139-149">Le code HTML utilise les composants de l’interface utilisateur Fluent React et spécifie que lorsque l’événement se déclenche, la fonction `Label` `PrimaryButton` `onClick` `insertText` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="33139-149">The HTML markup uses the `Label` and `PrimaryButton` components from Fluent UI React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="33139-150">Ajoutez le composant React à votre complément</span><span class="sxs-lookup"><span data-stu-id="33139-150">Add the React component to your add-in</span></span>

<span data-ttu-id="33139-151">Ajoutez le composant `ButtonPrimaryExample` à votre complément en ouvrant **src\components\App.tsx** et en effectuant les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="33139-151">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="33139-152">Ajoutez l’instruction importation suivante pour référencer `ButtonPrimaryExample` dans **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="33139-152">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="33139-153">Supprimez les deux instructions d’importation suivantes.</span><span class="sxs-lookup"><span data-stu-id="33139-153">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="33139-154">Remplacez la fonction `render()` par défaut par le code suivant qui utilise `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="33139-154">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

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

4. <span data-ttu-id="33139-155">Enregistrez les modifications apportées à **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="33139-155">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="33139-156">Regardez le résultat</span><span class="sxs-lookup"><span data-stu-id="33139-156">See the result</span></span>

<span data-ttu-id="33139-157">Dans Word, le volet Office complément se met automatiquement à jour lorsque vous enregistrez les modifications apportées à **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="33139-157">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="33139-158">Le texte et le bouton par défaut en bas du volet Office indiquent désormais l’interface utilisateur définie par le composant `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="33139-158">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="33139-159">Sélectionnez le bouton **Insérer un texte...** pour insérer du texte dans le document.</span><span class="sxs-lookup"><span data-stu-id="33139-159">Choose the **Insert text...** button to insert text into the document.</span></span>

![Capture d’écran montrant l’application Word avec « Insérer du texte... » bouton et texte qui précède immédiatement mis en surbrill](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="33139-161">Félicitations, vous avez créé un add-in du volet Des tâches à l’aide de React’interface utilisateur Fluent et React !</span><span class="sxs-lookup"><span data-stu-id="33139-161">Congratulations, you've successfully created a task pane add-in using React and Fluent UI React!</span></span>

## <a name="see-also"></a><span data-ttu-id="33139-162">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="33139-162">See also</span></span>

- [<span data-ttu-id="33139-163">Word Add-in GettingStartedFabricReact</span><span class="sxs-lookup"><span data-stu-id="33139-163">Word Add-in GettingStartedFabricReact</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="33139-164">Fabric Core dans les Office de base</span><span class="sxs-lookup"><span data-stu-id="33139-164">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="33139-165">Modèles de conception de l’expérience utilisateur pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="33139-165">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
