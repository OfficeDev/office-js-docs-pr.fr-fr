---
title: Utilisation d’Office UI Fabric React dans des compléments Office
description: Découvrez comment utiliser Office UI Fabric React dans les compléments Office.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: f8f61d1b094fa71b8a400a6a6d9ea3029c53b051
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237727"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="70b8e-103">Utilisation d’Office UI Fabric React dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="70b8e-103">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="70b8e-p101">Office UI Fabric est une infrastructure frontale JavaScript pour créer des expériences utilisateur pour Office. Si vous créez votre add-in à l’aide de React, envisagez d’utiliser Fabric React pour créer votre expérience utilisateur. Fabric fournit plusieurs composants UX react, tels que des boutons ou des case à cocher, que vous pouvez utiliser dans votre add-in.</span><span class="sxs-lookup"><span data-stu-id="70b8e-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="70b8e-107">Cet article décrit la création d’un complément conçu avec la fonction React et utilise les composants Fabric React.</span><span class="sxs-lookup"><span data-stu-id="70b8e-107">This article describes how to create an add-in that's built with React and uses Fabric React components.</span></span>

> [!NOTE]
> <span data-ttu-id="70b8e-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) est inclus dans Fabric React, ce qui signifie que votre complément aura également accès à Fabric Core une fois que vous aurez effectué les étapes décrites dans cet article.</span><span class="sxs-lookup"><span data-stu-id="70b8e-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) is included with Fabric React, which means your add-in will also have access to Fabric Core after you've completed the steps in this article.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="70b8e-109">Création d’un projet de complément</span><span class="sxs-lookup"><span data-stu-id="70b8e-109">Create an add-in project</span></span>

<span data-ttu-id="70b8e-110">Vous utiliserez le générateur Yeoman pour les compléments Office pour créer un projet de complément utilisant React.</span><span class="sxs-lookup"><span data-stu-id="70b8e-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="70b8e-111">Installez les composants requis</span><span class="sxs-lookup"><span data-stu-id="70b8e-111">Install the prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="70b8e-112">Créez le projet</span><span class="sxs-lookup"><span data-stu-id="70b8e-112">Create the project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="70b8e-113">**Sélectionnez un type de projet :** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="70b8e-113">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="70b8e-114">**Sélectionnez un type de script :** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="70b8e-114">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="70b8e-115">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="70b8e-115">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="70b8e-116">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="70b8e-116">**Which Office client application would you like to support?**</span></span> `Word`

![Capture d’écran montrant les invites et réponses relatives au générateur Yeoman dans une interface de ligne de commande](../images/yo-office-word-react.png)

<span data-ttu-id="70b8e-118">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="70b8e-118">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a><span data-ttu-id="70b8e-119">Essayez</span><span class="sxs-lookup"><span data-stu-id="70b8e-119">Try it out</span></span>

1. <span data-ttu-id="70b8e-120">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="70b8e-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="70b8e-121">Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="70b8e-121">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="70b8e-122">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="70b8e-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="70b8e-123">Si vous êtes invité à installer un certificat après avoir exécuté une des commandes suivantes, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="70b8e-123">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="70b8e-124">Il se peut également que vous deviez exécuter votre invite de commande ou votre terminal en tant qu'administrateur pour que les modifications soient effectuées.</span><span class="sxs-lookup"><span data-stu-id="70b8e-124">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    > [!TIP]
    > <span data-ttu-id="70b8e-125">Si vous testez votre complément sur Mac, exécutez la commande suivante avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="70b8e-125">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="70b8e-126">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="70b8e-126">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="70b8e-127">Pour tester votre complément dans Word, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="70b8e-127">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="70b8e-128">Cela a pour effet de démarrer le serveur web local (s’il n’est pas déjà en cours d’exécution) et d’ouvrir Word avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="70b8e-128">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="70b8e-129">Pour tester votre complément dans Word sur un navigateur, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="70b8e-129">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="70b8e-130">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="70b8e-130">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="70b8e-131">Pour utiliser votre complément, ouvrez un nouveau document dans Word sur le web, puis chargez la version test de votre complément en suivant les instructions de l’article relatif au [chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="70b8e-131">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="70b8e-132">Dans Word, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="70b8e-132">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="70b8e-133">Remarquez le texte par défaut et le bouton **Exécuter** en bas du volet Office.</span><span class="sxs-lookup"><span data-stu-id="70b8e-133">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="70b8e-134">Ensuite, vous redéfinirez ce texte et ce bouton en créant un composant React qui utilise les composants UX de Fabric React.</span><span class="sxs-lookup"><span data-stu-id="70b8e-134">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fabric React.</span></span>

    ![Screenshot showing the Word application with the Show Taskpane ribbon button highlighted and the Run button and immediately preceding text highlighted in the task pane](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fabric-react"></a><span data-ttu-id="70b8e-136">Créer un composant React utilisant Fabric React</span><span class="sxs-lookup"><span data-stu-id="70b8e-136">Create a React component that uses Fabric React</span></span>

<span data-ttu-id="70b8e-137">À ce stade, vous avez créé un complément très rudimentaire du volet Office standard en utilisant React.</span><span class="sxs-lookup"><span data-stu-id="70b8e-137">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="70b8e-138">Ensuite, procédez comme suit pour créer un nouveau composant React (`ButtonPrimaryExample`) dans le projet de complément.</span><span class="sxs-lookup"><span data-stu-id="70b8e-138">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="70b8e-139">Le composant utilise les composants `Label` et `PrimaryButton` de Fabric React.</span><span class="sxs-lookup"><span data-stu-id="70b8e-139">The component uses the `Label` and `PrimaryButton` components from Fabric React.</span></span>

1. <span data-ttu-id="70b8e-140">Ouvrez le dossier du projet créé par le générateur Yeoman et accédez à **src\taskpane\components**.</span><span class="sxs-lookup"><span data-stu-id="70b8e-140">Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.</span></span>
2. <span data-ttu-id="70b8e-141">Dans ce dossier, créez un fichier nommé **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="70b8e-141">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="70b8e-142">Dans **Button.tsx**, ajoutez le code suivant pour définir le composant `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="70b8e-142">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

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
      body.insertParagraph('Hello Office UI Fabric React!', Word.InsertLocation.end);
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

<span data-ttu-id="70b8e-143">Ce code effectue les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="70b8e-143">This code does the following:</span></span>

- <span data-ttu-id="70b8e-144">Fait référence à la bibliothèque React en utilisant `import * as React from 'react';`.</span><span class="sxs-lookup"><span data-stu-id="70b8e-144">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="70b8e-145">Référence les composants de Fabric (`PrimaryButton`, `IButtonProps`, `Label`) utilisés pour créer `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="70b8e-145">References the Fabric components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="70b8e-146">Déclare le nouveau composant `ButtonPrimaryExample` en utilisant `export class ButtonPrimaryExample extends React.Component`.</span><span class="sxs-lookup"><span data-stu-id="70b8e-146">Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="70b8e-147">Déclare la fonction `insertText` qui gère l’événement du bouton `onClick`.</span><span class="sxs-lookup"><span data-stu-id="70b8e-147">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="70b8e-148">Définit l’interface utilisateur du composant React dans la fonction `render`.</span><span class="sxs-lookup"><span data-stu-id="70b8e-148">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="70b8e-149">Le balisage HTML utilise les composants `Label` et `PrimaryButton` de Fabric React et spécifie que lorsque l’événement `onClick` se déclenche, la fonction `insertText` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="70b8e-149">The HTML markup uses the `Label` and `PrimaryButton` components from Fabric React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="70b8e-150">Ajoutez le composant React à votre complément</span><span class="sxs-lookup"><span data-stu-id="70b8e-150">Add the React component to your add-in</span></span>

<span data-ttu-id="70b8e-151">Ajoutez le composant `ButtonPrimaryExample` à votre complément en ouvrant **src\components\App.tsx** et en effectuant les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="70b8e-151">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="70b8e-152">Ajoutez l’instruction importation suivante pour référencer `ButtonPrimaryExample` dans **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="70b8e-152">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="70b8e-153">Supprimez les deux instructions d’importation suivantes.</span><span class="sxs-lookup"><span data-stu-id="70b8e-153">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="70b8e-154">Remplacez la fonction `render()` par défaut par le code suivant qui utilise `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="70b8e-154">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

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

4. <span data-ttu-id="70b8e-155">Enregistrez les modifications apportées à **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="70b8e-155">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="70b8e-156">Regardez le résultat</span><span class="sxs-lookup"><span data-stu-id="70b8e-156">See the result</span></span>

<span data-ttu-id="70b8e-157">Dans Word, le volet Office complément se met automatiquement à jour lorsque vous enregistrez les modifications apportées à **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="70b8e-157">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="70b8e-158">Le texte et le bouton par défaut en bas du volet Office indiquent désormais l’interface utilisateur définie par le composant `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="70b8e-158">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="70b8e-159">Sélectionnez le bouton **Insérer un texte...** pour insérer du texte dans le document.</span><span class="sxs-lookup"><span data-stu-id="70b8e-159">Choose the **Insert text...** button to insert text into the document.</span></span>

![Capture d’écran montrant l’application Word avec « Insérer du texte... » bouton et texte qui précède immédiatement mis en surbrill](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="70b8e-161">Félicitations, vous avez créé un complément de volet Office à l’aide de React et Office UI Fabric React !</span><span class="sxs-lookup"><span data-stu-id="70b8e-161">Congratulations, you've successfully created a task pane add-in using React and Office UI Fabric React!</span></span>

## <a name="see-also"></a><span data-ttu-id="70b8e-162">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="70b8e-162">See also</span></span>

- [<span data-ttu-id="70b8e-163">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="70b8e-163">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
- [<span data-ttu-id="70b8e-164">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="70b8e-164">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="70b8e-165">Modèles de conception de l’expérience utilisateur pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="70b8e-165">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
- [<span data-ttu-id="70b8e-166">Démarrer avec un exemple de code Fabric React</span><span class="sxs-lookup"><span data-stu-id="70b8e-166">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
