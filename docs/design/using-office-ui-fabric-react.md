---
title: Utilisation d’Office UI Fabric React dans des compléments Office
description: ''
ms.date: 2/28/2019
localization_priority: Priority
ms.openlocfilehash: 7d3e280298ee6761be9e7ced96d3490defeef7f0
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359239"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="8590d-102">Utilisation d’Office UI Fabric React dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="8590d-102">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="8590d-p101">Office UI Fabric est l’infrastructure frontale JavaScript permettant de créer des expériences pour Office et Office 365. Si vous créez votre complément à l’aide de React, envisagez d’utiliser Fabric React pour créer votre expérience utilisateur. Fabric fournit plusieurs composants UX basés sur React, tels que des boutons ou cases à cocher, que vous pouvez utiliser dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="8590d-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="8590d-106">Pour commencer à utiliser les composants de Fabric React dans votre complément, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="8590d-106">To get started using Fabric React's components in your add-in, perform the following steps.</span></span>

> [!NOTE]
> <span data-ttu-id="8590d-107">Si vous suivez les étapes de cet article, Fabric Core est également disponible dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="8590d-107">If you follow the steps in this article, Fabric Core is also available in your add-in.</span></span>

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a><span data-ttu-id="8590d-108">Étape 1 : créez votre projet avec le générateur Yeoman pour Office</span><span class="sxs-lookup"><span data-stu-id="8590d-108">Step 1 - Create your project with the Yeoman generator for Office</span></span>

<span data-ttu-id="8590d-109">Pour créer un complément qui utilise Fabric React, nous recommandons d’utiliser le générateur Yeoman pour Office.</span><span class="sxs-lookup"><span data-stu-id="8590d-109">To create an add-in that uses Fabric React, we recommend that you use the Yeoman generator for Office.</span></span> <span data-ttu-id="8590d-110">Le générateur Yeoman pour Office fournit la génération automatique de modèles de projet et la gestion de création nécessaires au développement d’un complément Office.</span><span class="sxs-lookup"><span data-stu-id="8590d-110">The Yeoman generator for Office provides the project scaffolding and build management needed to develop an Office Add-in.</span></span>

<span data-ttu-id="8590d-111">Pour créer votre projet, procédez comme suit à l’aide de **Windows PowerShell** (pas l’invite de commande) :</span><span class="sxs-lookup"><span data-stu-id="8590d-111">To create your project, perform the following steps using **Windows PowerShell** (not the command prompt):</span></span>

1. <span data-ttu-id="8590d-112">Installez les éléments prérequis.</span><span class="sxs-lookup"><span data-stu-id="8590d-112">Install the prerequisites.</span></span>
2. <span data-ttu-id="8590d-113">Exécutez `yo office` pour créer les fichiers de projet pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="8590d-113">Run `yo office` to create the project files for your add-in.</span></span>
3. <span data-ttu-id="8590d-114">Lorsque vous êtes invité à sélectionner une application client Office, choisissez **Word**.</span><span class="sxs-lookup"><span data-stu-id="8590d-114">When prompted to select an Office client application, choose **Word**.</span></span>
4. <span data-ttu-id="8590d-p103">Vérifiez que vous êtes dans le répertoire contenant les fichiers de projet, puis exécutez `npm start`. Une fenêtre du navigateur affichant un bouton fléché s’ouvre automatiquement.</span><span class="sxs-lookup"><span data-stu-id="8590d-p103">Ensure you are in the directory with the project files, and then run `npm start`. A browser window showing a spinner opens automatically.</span></span>
5. <span data-ttu-id="8590d-117">[Chargez une version test de votre manifeste](..\testing\test-debug-office-add-ins.md) pour afficher l’interface utilisateur complète du complément.</span><span class="sxs-lookup"><span data-stu-id="8590d-117">[Sideload your manifest](..\testing\test-debug-office-add-ins.md) to view the full UI of the add-in.</span></span>

## <a name="step-2---add-a-fabric-react-component"></a><span data-ttu-id="8590d-118">Étape 2 : ajoutez un composant Fabric React</span><span class="sxs-lookup"><span data-stu-id="8590d-118">Step 2 - Add a Fabric React component</span></span>

<span data-ttu-id="8590d-p104">Ensuite, ajoutez des composants Fabric React à votre complément. Créez un nouveau composant REACT, appelé `ButtonPrimaryExample`, constitué d’une étiquette et d’un PrimaryButton de Fabric React. Pour créer `ButtonPrimaryExample` :</span><span class="sxs-lookup"><span data-stu-id="8590d-p104">Next, add Fabric React components to your add-in. Create a new React component, called `ButtonPrimaryExample`, that consists of a Label and PrimaryButton from Fabric React. To create `ButtonPrimaryExample`:</span></span>

1. <span data-ttu-id="8590d-122">Ouvrez le dossier du projet créé par le générateur Yeoman et accédez à **src\components**.</span><span class="sxs-lookup"><span data-stu-id="8590d-122">Open the project folder created by the Yeoman generator, and go to **src\components**.</span></span>
2. <span data-ttu-id="8590d-123">Créez **button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="8590d-123">Create **button.tsx**.</span></span>
3. <span data-ttu-id="8590d-124">Dans **button.tsx**, entrez le code suivant pour créer le composant `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="8590d-124">In **button.tsx**, enter the following code to create the `ButtonPrimaryExample` component.</span></span>

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor() {
    super();
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

<span data-ttu-id="8590d-125">Ce code effectue les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="8590d-125">This code does the following:</span></span>

- <span data-ttu-id="8590d-126">Fait référence à la bibliothèque React en utilisant `import * as React from 'react';`.</span><span class="sxs-lookup"><span data-stu-id="8590d-126">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="8590d-127">Fait référence aux composants Fabric (PrimaryButton, IButtonProps, étiquette) qui sont utilisés pour créer `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="8590d-127">References the Fabric components (PrimaryButton, IButtonProps, Label) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="8590d-128">Déclare et publie le nouveau composant `ButtonPrimaryExample` à l’aide de `export class ButtonPrimaryExample extends React.Component`.</span><span class="sxs-lookup"><span data-stu-id="8590d-128">Declares and make public the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="8590d-129">Déclare la fonction `insertText` pour gérer l’événement `onClick`.</span><span class="sxs-lookup"><span data-stu-id="8590d-129">Declares the `insertText` function to handle the `onClick` event.</span></span>
- <span data-ttu-id="8590d-p105">Définit l’interface utilisateur du composant React dans la fonction `render`. Cette fonction définit la structure du composant. Dans `render`, vous associez l’événement `this.insertText` en utilisant `onClick`.</span><span class="sxs-lookup"><span data-stu-id="8590d-p105">Defines the UI of the React component in the `render` function. Render defines the structure of the component. Within `render`, you wire up the `onClick` event using `this.insertText`.</span></span>

## <a name="step-3---add-the-react-component-to-your-add-in"></a><span data-ttu-id="8590d-133">Étape 3 : ajoutez le composant React à votre complément</span><span class="sxs-lookup"><span data-stu-id="8590d-133">Step 3 - Add the React component to your add-in</span></span>

<span data-ttu-id="8590d-134">Ajoutez `ButtonPrimaryExample` à votre complément en ouvrant **src\components\app.tsx** et en effectuant les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="8590d-134">Add `ButtonPrimaryExample` to your add-in by opening **src\components\app.tsx** and doing the following:</span></span>

- <span data-ttu-id="8590d-135">Ajoutez l’instruction d’importation suivante pour faire référence à `ButtonPrimaryExample` depuis le **button.tsx** créé à l’étape 2 (aucune extension de fichier n’est nécessaire).</span><span class="sxs-lookup"><span data-stu-id="8590d-135">Add the following import statement to reference `ButtonPrimaryExample` from **button.tsx** created in step 2 (no file extension is needed).</span></span>

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- <span data-ttu-id="8590d-136">Remplacez la fonction `render()` par défaut par le code suivant qui utilise `<ButtonPrimaryExample />`.</span><span class="sxs-lookup"><span data-stu-id="8590d-136">Replace the default `render()` function with the following code that uses `<ButtonPrimaryExample />`.</span></span>

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

<span data-ttu-id="8590d-p106">Enregistrez vos modifications. Toutes les instances de navigateur ouvertes, y compris le complément, sont mises à jour automatiquement et affichent le composant React `ButtonPrimaryExample`. Vous pouvez remarquer que le texte par défaut et le bouton sont remplacés par le texte et le bouton principal définis dans `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="8590d-p106">Save your changes. All open browser instances, including the add-in, update automatically and show the `ButtonPrimaryExample` React component. Notice that the default text and button is replaced with the text and primary button defined in `ButtonPrimaryExample`.</span></span>



## <a name="see-also"></a><span data-ttu-id="8590d-140">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8590d-140">See also</span></span>

- [<span data-ttu-id="8590d-141">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="8590d-141">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="8590d-142">Modèles de conception de l’expérience utilisateur pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="8590d-142">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
- [<span data-ttu-id="8590d-143">Démarrer avec un exemple de code Fabric React</span><span class="sxs-lookup"><span data-stu-id="8590d-143">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="8590d-144">Exemples d’éléments d’interface utilisateur Fabric pour les compléments Office (utilise Fabric 1.0)</span><span class="sxs-lookup"><span data-stu-id="8590d-144">Office Add-in Fabric UI sample (uses Fabric 1.0)</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="8590d-145">Générateur Yeoman pour Office</span><span class="sxs-lookup"><span data-stu-id="8590d-145">Yeoman generator for Office</span></span>](https://github.com/OfficeDev/generator-office)
