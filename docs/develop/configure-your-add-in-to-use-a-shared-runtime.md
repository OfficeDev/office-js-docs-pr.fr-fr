---
ms.date: 10/05/2021
title: Configurez votre complément Office pour utiliser un runtime JavaScript partagé
ms.prod: non-product-specific
description: Configurez votre complément Office afin d’utiliser un runtime JavaScript partagé pour prendre en charge un ruban supplémentaire, un volet des tâches et des fonctionnalités personnalisées.
ms.localizationpriority: high
ms.openlocfilehash: 95a4cb410bf92a68c1790e3fba67ea482bdc78b6
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138471"
---
# <a name="configure-your-office-add-in-to-use-a-shared-javascript-runtime"></a>Configurez votre complément Office pour utiliser un runtime JavaScript partagé

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Vous pouvez configurer votre complément Office pour exécuter la totalité de son code dans un seul runtime JavaScript partagé (également connu sous le nom de runtime partagé). Vous pouvez ainsi améliorer la coordination dans votre complément et accéder aux DOM et CORS à partir de toutes les parties de votre complément. Il active également des fonctionnalités supplémentaires telles que l’exécution d’un code lors de l’ouverture d’un document, ou l’activation et la désactivation des boutons du ruban. Si vous voulez configurer votre complément pour utiliser un runtime partagé JavaScript, suivez les instructions contenues dans cet article.

## <a name="create-the-add-in-project"></a>Création du projet de complément

Si vous démarrez un nouveau projet, suivez ces étapes pour utiliser le [générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office) et créer le projet de complément Excel ou PowerPoint.

Effectuez l'une des opérations suivantes :

- Pour créer un complément Excel avec fonctions personnalisées, exécutez la commande `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true`.

    ou

- Pour créer un complément PowerPoint, exécutez la commande `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js true`.

Le générateur crée le projet et installe les composants de nœud de la prise en charge.

> [!NOTE]
> Vous pouvez également utiliser les étapes décrites dans cet article pour mettre à jour un projet Visual Studio existant afin d’utiliser le runtime partagé. Toutefois, vous devrez peut-être mettre à jour les schémas XML pour le manifeste. Pour plus d’informations, consultez [Résoudre les erreurs de développement avec les compléments Office](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

## <a name="configure-the-manifest"></a>Configurer le manifeste

Procédez comme suit pour configurer un projet nouveau ou existant de manière à utiliser un runtime partagé. Ces étapes supposent que vous avez créé votre projet à l’aide du [générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office).

1. Démarrez Visual Studio Code, puis ouvrez le projet de complément Excel ou PowerPoint créé.
1. Ouvrez le fichier **manifest.xml**.
1. Si vous avez généré un complément Excel, mettez à jour la section des exigences pour utiliser le [ de runtime partagé](../reference/requirement-sets/shared-runtime-requirement-sets.md)au lieu du runtime de fonction personnalisé. Le code XML doit apparaître comme suit.

    ```xml
    <Hosts>
      <Host Name="Workbook"/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. Recherchez la `<VersionOverrides>` section et ajoutez la section `<Runtimes>` suivante. La durée de vie doit être **longue** afin que votre code de complément puisse s’exécuter même quand le volet Office est fermé. La valeur `resid` est **Taskpane.Url** qui se réfère à l’emplacement du fichier **taskpane.html** spécifiée dans la section ` <bt:Urls>` près du bas du fichier **manifest.xml**.

    > [!IMPORTANT]
    > La `<Runtimes>` section doit être entrée après `<Host>` l’élément dans l’ordre exact indiqué dans le XML suivant.

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
         <Runtimes>
           <Runtime resid="Taskpane.Url" lifetime="long" />
         </Runtimes>
       ...
       </Host>
   ```

1. Si vous avez créé un complément Excel avec des fonctions personnalisées, recherchez l’élément `<Page>`. Puis remplacez l’emplacement de la source **Functions.Page.Url** par **TaskPane.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. Recherchez la balise `<FunctionFile ...>` et remplacez le `resid` de **Commands.Url** par **Taskpane.Url**. Veuillez noter que si vous n'avez pas de commandes d'action, vous ne disposerez pas d'entrée **FunctionFile**. Vous pouvez par conséquent ignorer cette étape.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Enregistrez le fichier **manifest.xml**.

## <a name="configure-the-webpackconfigjs-file"></a>Configurer le fichier webpack.config.js

Le fichier **webpack.config.js** générera plusieurs chargeurs runtime. Vous devez le modifier pour charger uniquement le runtime JavaScript partagé via le fichier **taskpane.html**.

1. Démarrez Visual Studio Code, puis ouvrez le projet de complément Excel ou PowerPoint créé.
1. Ouvrez le fichier **webpack.config.js**.
1. Si votre fichier **webpack.config.js** a le code plug-in **functions.html**, supprimez-le.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. Si votre fichier **webpack.config.js** a le code plug-in **commands.html**, supprimez-le.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. Si votre projet utilisait les blocs **fonctions** ou **commandes**, ajoutez-les à la liste des blocs comme illustré par la suite (le code suivant sert si votre projet utilisait les deux blocs).

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. Enregistrez vos changements et reconstruisez le projet.

   ```command line
   npm run build
   ```

> [!NOTE]
> Si votre projet a le fichier **functions.html** ou le fichier **commands.html**, vous pouvez les supprimer. Le fichier **taskpane.html** chargera le code **functions.js** et **commands.js** dans le runtime JavaScript partagé via les mises à jour webpack que vous venez d’effectuer.

## <a name="test-your-office-add-in-changes"></a>Tester les modifications apportées à votre complément Office

Vous pouvez confirmer que vous utilisez correctement le runtime JavaScript partagé en utilisant les instructions suivantes.

1. Ouvrez le fichier **manifest.xml**.
1. Recherchez la section `<Control xsi:type="Button" id="TaskpaneButton">`, puis modifiez le XML `<Action ...>` suivant.

    de :

    ```xml
    <Action xsi:type="ShowTaskpane">
      <TaskpaneId>ButtonId1</TaskpaneId>
      <SourceLocation resid="Taskpane.Url"/>
    </Action>
    ```

    à :

    ```xml
    <Action xsi:type="ExecuteFunction">
      <FunctionName>action</FunctionName>
    </Action>
    ```

1. Ouvrez le fichier **./src/commands/commands.js**.
1. Remplacez la fonction **action** existante par le code suivant. Cette action mettra à jour la fonction pour ouvrir et modifier le bouton de volet des tâches pour incrémenter un compteur. L’ouverture et l’accès au volet des tâches DOM à partir d’une commande ne fonctionne qu’avec le runtime JavaScript partagé.

    ```javascript
    var _count=0;
    
    function action(event) {
      // Your code goes here.
      _count++;
      Office.addin.showAsTaskpane();
      document.getElementById("run").textContent="Go"+_count;
    
      // Be sure to indicate when the add-in command function is complete.
      event.completed();
    }
    ```

1. Enregistrez vos changements et exécutez le projet.

   ```command line
   npm start
   ```

Chaque fois que vous sélectionnez le bouton de complément, il changera le texte du bouton **exécuter** par **accéder** et incrémentera un compteur après celui-ci.

## <a name="runtime-lifetime"></a>Durée de vie de l’exécution

Lorsque vous ajoutez l’élément `Runtime`, vous spécifiez également une durée de vie avec la valeur `long` ou `short`. Définissez cette valeur sur `long` pour tirer parti de fonctionnalités telles que le démarrage de votre complément lorsque le document s’ouvre, la poursuite de l’exécution du code après la fermeture du volet Office ou l’utilisation de CORS et DOM à partir de fonctions personnalisées.

> [!NOTE]
> La valeur de la durée de vie par défaut est `short`, mais nous vous recommandons d’utiliser `long` dans les compléments Excel. Si vous avez défini votre runtime sur `short` dans cet exemple, votre complément Excel démarre lorsque vous appuyez sur l’un de vos boutons du ruban, mais il se peut qu’il se ferme une fois l’exécution de votre gestionnaire de ruban terminée. De la même façon, le complément démarre lorsque le volet des tâches est ouvert, mais il se peut se fermer à la fermeture du volet des tâches.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> Si votre macro complémentaire inclut l’`Runtimes`élément dans le manifeste (runtime partagé requis) et que les conditions d’utilisation de Microsoft Edge avec WebView2 (basées sur Chromium) sont remplies, il utilise ce contrôle WebView2. Si les conditions ne sont pas remplies, il utilise Internet Explorer 11, quelle que soit la version Windows ou Microsoft 365 version. Pour plus d’informations, consultez [Runtimes](../reference/manifest/runtimes.md) and [Browsers utilisés par les compléments Office ](../concepts/browsers-used-by-office-web-add-ins.md).

## <a name="about-the-shared-javascript-runtime"></a>À propos du runtime JavaScript partagé

Sur Windows ou Mac, votre complément exécute le code des boutons du ruban, des fonctions personnalisées et du volet des tâches dans des environnements runtime JavaScript distincts. Cela permet de créer des limitations, telles que l'impossibilité de partager aisément des données globales ou de pouvoir accéder à l'ensemble des fonctionnalités CORS à partir d’une fonction personnalisée.

Vous pouvez toutefois configurer votre complément Office pour partager un code dans le même runtime JavaScript (également appelé runtime partagé). Vous pouvez ainsi améliorer la coordination dans votre complément et accéder au volet des tâches DOM et CORS à partir de toutes les parties de votre complément.

La configuration d’un runtime partagé permet les scénarios suivants.

- Votre complément Office peut utiliser des fonctionnalités d’interface utilisateur supplémentaires :
  - [Ajouter des raccourcis clavier personnalisés à votre complément Office (préversion)](../design/keyboard-shortcuts.md)
  - [Créer des onglets contextuels personnalisés dans des compléments Office (préversion)](../design/contextual-tabs.md)
  - [Activer et désactiver des commandes de complément](../design/disable-add-in-commands.md)
  - [Exécuter un cote dans votre complément Office lors de l’ouverture du document](run-code-on-document-open.md)
  - [Afficher ou masquer le volet des tâches de votre complément Office](show-hide-add-in.md)
- Pour les compléments Excel :
  - Les fonctions personnalisées bénéficieront d'une prise en charge complète de CORS.
  - Les fonctions personnalisées peuvent appeler les API Office.js pour lire les données d’un document feuille de calcul.

Pour Office sur Windows, le runtime partagé utilise Microsoft Edge avec WebView2 (basé sur Chromium) si les conditions de son utilisation sont remplies comme expliqué dans [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md). Sinon, il utilise Internet Explorer 11. De plus, tous les boutons affichés par votre complément sur le ruban s’exécutent dans le même runtime partagé. L’image ci-après présente l'exécution des fonctions personnalisées, de interface utilisateur du ruban et du code du volet des tâches dans le même runtime JavaScript.

![Diagramme d’une fonction personnalisée, du volet des tâches et des boutons du ruban qui s’exécutent tous dans un runtime de navigateur partagé dans Excel.](../images/custom-functions-in-browser-runtime.png)

### <a name="debugging"></a>Débogage

Lors de l’utilisation d’un runtime partagé, vous ne pouvez pas utiliser Visual Studio Code pour déboguer des fonctions personnalisées dans Excel sur Windows à cette date. Vous devez utiliser les outils de développement à la place. Pour plus d'informations, voir le [Débogage des compléments avec les outils de développement sur Windows](../testing/debug-add-ins-using-f12-developer-tools-on-windows.md).

### <a name="multiple-task-panes"></a>Multiples volets des tâches

Ne concevez pas votre complément pour utiliser plusieurs volets des tâches si vous envisagez d’utiliser un runtime partagé. Un runtime partagé prend uniquement en charge l’utilisation d’un volet des tâches. Notez que tout volet des tâches sans `<TaskpaneID>` est considéré comme un volet des tâches différent.

## <a name="give-us-feedback"></a>Faites-nous part de vos commentaires

Nous aimerions connaître vos commentaires sur cette fonctionnalité. Si vous trouvez des bogues, des problèmes ou des demandes sur cette fonctionnalité, faites-le nous savoir en créant un problème GitHub dans le dépôt [office-js](https://github.com/OfficeDev/office-js).

## <a name="see-also"></a>Voir aussi

- [Appeler des API Excel à partir d'une fonction personnalisée](../excel/call-excel-apis-from-custom-function.md)
- [Ajouter des raccourcis clavier personnalisés à votre complément Office (préversion)](../design/keyboard-shortcuts.md)
- [Créer des onglets contextuels personnalisés dans des compléments Office (préversion)](../design/contextual-tabs.md)
- [Activer et désactiver des commandes de complément](../design/disable-add-in-commands.md)
- [Exécuter un cote dans votre complément Office lors de l’ouverture du document](run-code-on-document-open.md)
- [Afficher ou masquer le volet des tâches de votre complément Office](show-hide-add-in.md)
- [Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
