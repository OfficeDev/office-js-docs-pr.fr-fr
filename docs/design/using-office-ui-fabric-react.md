---
title: "Utilisation d’Office UI Fabric React dans des compléments\_Office"
description: ''
ms.date: 12/04/2017
---

# <a name="use-office-ui-fabric-react-in-office-add-ins"></a>Utilisation d’Office UI Fabric React dans des compléments Office

Office UI Fabric est l’infrastructure frontale JavaScript permettant de créer des expériences pour Office et Office 365. Si vous créez votre complément à l’aide de React, envisagez d’utiliser Fabric React pour créer votre expérience utilisateur. Fabric fournit plusieurs composants UX basés sur React, tels que des boutons ou cases à cocher, que vous pouvez utiliser dans votre complément. 

Pour commencer à utiliser les composants de Fabric React dans votre complément, procédez comme suit.

> [!NOTE]
> Si vous suivez les étapes de cet article, Fabric Core est également disponible dans votre complément.

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a>Étape 1 : créez votre projet avec le générateur Yeoman pour Office

Pour créer un complément qui utilise Fabric React, nous recommandons d’utiliser le générateur Yeoman pour Office. Le générateur Yeoman pour Office fournit la génération automatique de modèles de projet et la gestion de création nécessaires au développement d’un complément Office. 

Pour créer votre projet, procédez comme suit à l’aide de **Windows PowerShell** (pas l’invite de commande) : 

1. Installez les éléments prérequis.
2. Exécutez `yo office` pour créer les fichiers de projet pour votre complément. 
3. Lorsque vous êtes invité à sélectionner une application client Office, choisissez **Word**. 
4. Vérifiez que vous êtes dans le répertoire contenant les fichiers de projet, puis exécutez `npm start`. Une fenêtre du navigateur affichant un bouton fléché s’ouvre automatiquement.
5. Chargez votre manifeste pour afficher l’interface utilisateur complète du complément.    

## <a name="step-2---add-a-fabric-react-component"></a>Étape 2 : ajoutez un composant Fabric React

Ensuite, ajoutez des composants Fabric React à votre complément. Créez un nouveau composant REACT, appelé `ButtonPrimaryExample`, constitué d’une étiquette et d’un PrimaryButton de Fabric React. Pour créer `ButtonPrimaryExample` :

1. Ouvrez le dossier du projet créé par le générateur Yeoman et accédez à **src\components**.
2. Créez **button.tsx**.
3. Dans **button.tsx**, entrez le code suivant pour créer le composant `ButtonPrimaryExample`. 

```javascript
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
            var body = context.document.body;  
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
          onClick={ this.insertText }
        />
      </div>
    );
  }
}
```

Ce code effectue les opérations suivantes :

- Fait référence à la bibliothèque React en utilisant `import * as React from 'react';`.
- Fait référence aux composants Fabric (PrimaryButton, IButtonProps, étiquette) qui sont utilisés pour créer `ButtonPrimaryExample`. 
- Déclare et publie le nouveau composant `ButtonPrimaryExample` à l’aide de `export class ButtonPrimaryExample extends React.Component`. 
- Déclare la fonction `insertText` pour gérer l’événement onclick. 
- Définit l’interface utilisateur du composant React dans la fonction `render`. Le rendu définit la structure du composant. Dans `render`, vous associez l’événement onclick en utilisant `this.insertText`.

## <a name="step-3---add-the-react-component-to-your-add-in"></a>Étape 3 : ajoutez le composant React à votre complément 

Ajoutez `ButtonPrimaryExample` à votre complément en ouvrant **src\components\app.tsx** et en effectuant les opérations suivantes : 

- Ajoutez l’instruction d’importation suivante pour faire référence à `ButtonPrimaryExample` depuis le **button.tsx** créé à l’étape 2 (aucune extension de fichier n’est nécessaire). 

    ```javascript
    import {ButtonPrimaryExample} from './button';
    ``` 

- Remplacez la fonction `render()` par défaut par le code suivant qui utilise `<ButtonPrimaryExample />`. 

  ```javascript
  render() {
      return (
          <div className='ms-welcome'>
              <Header logo='assets/logo-filled.png' title={this.props.title} message='Welcome' />
              <HeroList message='Discover what this add-in can do for you today!' items={this.state.listItems}>                    
                  <ButtonPrimaryExample />
              </HeroList>
          </div>
      );
  };
  ```

Enregistrez vos modifications. Toutes les instances de navigateur ouvertes, y compris le complément, sont mises à jour automatiquement et affichent le composant React `ButtonPrimaryExample`. Vous pouvez remarquer que le texte par défaut et le bouton sont remplacés par le texte et le bouton principal définis dans `ButtonPrimaryExample`. 
    
## <a name="recommended-components"></a>Composants recommandés

Voici une liste des composants UX Fabric React que nous vous recommandons d’utiliser dans un complément.  

> [!NOTE]
> nous allons ajouter des composants supplémentaires au fil du temps. 

- [Barre de navigation](breadcrumb.md)
- [Bouton](button.md)
- [Case à cocher](checkbox.md)
- [ChoiceGroup](choicegroup.md)
- [Liste déroulante](dropdown.md)
- [Étiquette](label.md)
- [Liste](list.md)
- [Tableau croisé dynamique](pivot.md)
- [TextField](textfield.md)
- [Bouton bascule](toggle.md)

## <a name="see-also"></a>Voir aussi

- [Office UI Fabric React](https://dev.office.com/fabric#/)
- [Démarrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [Modèles de conception de l’expérience utilisateur (utilise Fabric 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Exemples d’éléments d’interface utilisateur Fabric pour les compléments Office (utilise Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [Utiliser Office UI Fabric 2.6.1 dans des compléments Office](ui-elements/using-office-ui-fabric.md)
- [Générateur Yeoman pour Office](https://github.com/OfficeDev/generator-office)
 

