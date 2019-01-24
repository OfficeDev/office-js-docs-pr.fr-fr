---
ms.openlocfilehash: f090023acbeae8bab9e25a36b64669c7d033daea
ms.sourcegitcommit: 319780e4fb77b97d16d5657950e4eb84415eb7eb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/17/2019
ms.locfileid: "28725527"
---
# <a name="build-an-excel-add-in-using-react"></a>Développement d’un complément Excel à l’aide de React

Cet article décrit le processus de création d’un complément Excel à l’aide de React et de l’API JavaScript pour Excel.

## <a name="prerequisites"></a>Conditions préalables

- [Node.js](https://nodejs.org)

- Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.
    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a>Création de l’application web

1. Utilisez le générateur Yeoman afin de créer un projet de complément Excel. Exécutez la commande suivante, puis répondez aux invites comme suit :

    ```bash
    yo office
    ```

    - **Sélectionnez un type de projet :** `Office Add-in project using React framework`
    - **Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ? :**`Excel`

    ![Générateur Yeoman](../images/yo-office-excel-react.png)
    
    Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants de nœud de la prise en charge.

2. Accédez au dossier racine du projet.

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a>Mise à jour du code

1. Dans votre éditeur de code, ouvrez le fichier **src/styles.less**, ajoutez les styles suivants à la fin du fichier et enregistrez-le.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
        font-family: Arial;
        padding-top: 25px;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
        font-family: Arial;
    }

    .padding {
        padding: 15px;
    }

    .padding-sm {
        padding: 4px;
    }

    .normal-button {
        width: 80px;
        padding: 2px;
    }
    ```

2. Le modèle de projet créés dans le générateur Yeoman de compléments Office inclut un composant React qui n’est pas nécessaire pour ce guide de démarrage rapide. Supprimez le fichier **src/components/HeroList.tsx**.

3. Ouvrez le fichier **src/components/Header.tsx**, remplacez-en tout le contenu par le code suivant et enregistrez le fichier.

    ```typescript
    import * as React from 'react';

    export interface HeaderProps {
        title: string;
    }

    export class Header extends React.Component<HeaderProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-header'>
                    <div className='padding'>
                        <h1>{this.props.title}</h1>
                    </div>
                </div>
            );
        }
    }
    ```

4. Créez un nouveau composant React nommé **Content.tsx** dans le dossier **src/composants**, ajoutez le code suivant et enregistrez le fichier.

    ```typescript
    import * as React from 'react';
    import { Button, ButtonType } from 'office-ui-fabric-react';

    export interface ContentProps {
        message: string;
        buttonLabel: string;
        click: any;
    }

    export class Content extends React.Component<ContentProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-main'>
                    <div className='padding'>
                        <p>{this.props.message}</p>
                        <br />
                        <h3>Try it out</h3>
                        <br/>
                        <Button className='normal-button' buttonType={ButtonType.hero} onClick={this.props.click}>{this.props.buttonLabel}</Button>
                    </div>
                </div>
            );
        }
    }
    ```

5. Ouvrez le fichier **src/components/App.tsx**, remplacez-en tout le contenu par le code suivant et enregistrez le fichier.

    ```typescript
    /* global Office, Excel */

    import * as React from 'react';
    import { Header } from './Header';
    import { Content } from './Content';
    import Progress from './Progress';

    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    export interface AppProps {
        title: string;
        isOfficeInitialized: boolean;
    }

    export interface AppState {
    }

    export default class App extends React.Component<AppProps, AppState> {
        constructor(props, context) {
            super(props, context);
        }

        setColor = async () => {
            try {
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    range.load('address');
                    range.format.fill.color = 'green';
                    await context.sync();
                    console.log(`The range address was ${range.address}.`);
                });
            } catch (error) {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            }
        }

        render() {
            const {
                title,
                isOfficeInitialized,
            } = this.props;

            if (!isOfficeInitialized) {
                return (
                    <Progress
                        title={title}
                        logo='assets/logo-filled.png'
                        message='Please sideload your addin to see app body.'
                    />
                );
            }

            return (
                <div className='ms-welcome'>
                    <Header title='Welcome' />
                    <Content message='Choose the button below to set the color of the selected range to green.' buttonLabel='Set color' click={this.setColor} />
                </div>
            );
        }
    }
    ```

## <a name="update-the-manifest"></a>Mise à jour du manifeste

1. Ouvrez le fichier nommé **manifest.xml** pour définir les paramètres et les fonctionnalités du complément. 

2. L’élément `ProviderName` possède une valeur d’espace réservé. Remplacez-le par votre nom.

3. L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé. Remplacez-le par **A task pane add-in for Excel**.

4. Enregistrez le fichier.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a>Démarrage du serveur de développement

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a>Essayez !

1. Suivez les instructions pour la plateforme que vous utiliserez afin d’exécuter votre complément en vue d’en charger une version test dans Excel.

    - Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)
    - iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2b.png)

3. Sélectionnez une plage de cellules dans la feuille de calcul.

4. Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément Excel à l’aide de React ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.

> [!div class="nextstepaction"]
> [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>Voir aussi

* [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial-create-table.md)
* [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](../excel/excel-add-ins-core-concepts.md)
* [Exemples de code pour les compléments Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Référence de l’API JavaScript pour Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
