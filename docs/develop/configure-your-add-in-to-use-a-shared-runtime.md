---
title: Configurer votre complément Office pour utiliser un runtime partagé
description: Configurez votre complément Office pour utiliser un runtime partagé pour prendre en charge les fonctionnalités supplémentaires du ruban, du volet Office et des fonctions personnalisées.
ms.date: 07/18/2022
ms.prod: non-product-specific
ms.localizationpriority: high
ms.openlocfilehash: e6b10cc2d342d95a8542146ecbd95d750322421f
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422935"
---
# <a name="configure-your-office-add-in-to-use-a-shared-runtime"></a>Configurer votre complément Office pour utiliser un runtime partagé

[!include[Shared runtime requirements](../includes/shared-runtime-requirements-note.md)]

Vous pouvez configurer votre complément Office pour exécuter tout son code dans un seul [runtime partagé](../testing/runtimes.md#shared-runtime). Vous pouvez ainsi améliorer la coordination dans votre complément et accéder aux DOM et CORS à partir de toutes les parties de votre complément. Il active également des fonctionnalités supplémentaires telles que l’exécution d’un code lors de l’ouverture d’un document, ou l’activation et la désactivation des boutons du ruban. Si vous voulez configurer votre complément pour utiliser un runtime partagé, suivez les instructions contenues dans cet article.

## <a name="create-the-add-in-project"></a>Création du projet de complément

Si vous démarrez un nouveau projet, utilisez le [Générateur Yeoman pour compléments Office](yeoman-generator-overview.md) pour créer le projet de complément Excel, PowerPoint ou Word.

Exécutez la commande `yo office --projectType taskpane --name "my office add in" --host <host> --js true`, où `<host>` est l’une des valeurs suivantes.

- excel
- powerpoint
- word

> [!IMPORTANT]
> La valeur de l’argument `--name` doit être entre guillemets doubles, même si elle n’a pas d’espace.

Vous pouvez utiliser différentes options pour les options de ligne de commande **--projecttype**, **--name** et **--js**. Pour obtenir la liste complète des options, consultez [Générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office).

Le générateur crée le projet et installe les composants de nœud de la prise en charge. Vous pouvez également utiliser les étapes décrites dans cet article pour mettre à jour un projet Visual afin d’utiliser le runtime partagé. Toutefois, vous devrez peut-être mettre à jour les schémas XML pour le manifeste. Pour plus d’informations, consultez [Résoudre les erreurs de développement avec les compléments Office](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

## <a name="configure-the-manifest"></a>Configurer le manifeste

Procédez comme suit pour configurer un projet nouveau ou existant de manière à utiliser un runtime partagé. Ces étapes supposent que vous avez créé votre projet à l’aide du [générateur Yeoman pour compléments Office](yeoman-generator-overview.md).

1. Démarrez Visual Studio Code, puis ouvrez votre projet de complément.
1. Ouvrez le fichier **manifest.xml**.
1. Pour un complément Excel ou PowerPoint, mettez à jour la section des conditions requises pour inclure le [runtime partagé](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets). Veillez à supprimer la condition requise `CustomFunctionsRuntime` si elle est présente. Le XML s’affiche comme suit.

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

    > [!NOTE]
    > N’ajoutez pas l’ensemble de conditions requises `SharedRuntime` au manifeste pour un complément Word. Cela génère une erreur lors du chargement du complément, qui est un problème connu pour l’instant.

1. Recherchez la section **\<VersionOverrides\>** et ajoutez la section **\<Runtimes\>** suivante. La durée de vie doit être **longue** afin que votre code de complément puisse s’exécuter même quand le volet Office est fermé. La valeur `resid` est **Taskpane.Url** qui se réfère à l’emplacement du fichier **taskpane.html** spécifiée dans la section `<bt:Urls>` près du bas du fichier **manifest.xml**.

    > [!IMPORTANT]
    > La section **\<Runtimes\>** doit être entrée après l’élément **\<Host\>** dans l’ordre exact indiqué dans le XML suivant.

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

1. Si vous avez créé un complément Excel avec des fonctions personnalisées, recherchez l’élément **\<Page\>**. Puis remplacez l’emplacement de la source **Functions.Page.Url** par **TaskPane.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. Recherchez la balise **\<FunctionFile\>**, puis remplacez le `resid` de **Commands.Url** par **Taskpane.Url**. Veuillez noter que si vous n'avez pas de commandes d'action, vous ne disposerez pas de l'entrée **\<FunctionFile\>**. Vous pouvez par conséquent ignorer cette étape.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Enregistrez le fichier **manifest.xml**.

## <a name="configure-the-webpackconfigjs-file"></a>Configurer le fichier webpack.config.js

Le fichier **webpack.config.js** générera plusieurs chargeurs runtime. Vous devez le modifier pour charger uniquement le runtime partagé via le fichier **taskpane.html** .

1. Démarrez Visual Studio Code et ouvrez le projet de complément que vous avez généré.
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
> Si votre projet a le fichier **functions.html** ou le fichier **commands.html**, vous pouvez les supprimer. Le **taskpane.html** charge le **codefunctions.js** et **commands.js** dans le runtime partagé via les mises à jour webpack que vous venez d’effectuer.

## <a name="test-your-office-add-in-changes"></a>Tester les modifications apportées à votre complément Office

Vous pouvez confirmer que vous utilisez correctement le runtime partagé en suivant les instructions suivantes.

1. Ouvrez le fichier **taskpane.js**.
1. Remplacez tout le contenu du fichier par le code suivant. Le nombre de fois où le volet Office a été ouvert s’affiche. L’ajout de l’événement onVisibilityModeChanged est uniquement pris en charge dans un runtime partagé.

    ```javascript
    /*global document, Office*/

    let _count = 0;

    Office.onReady(() => {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";

      updateCount(); // Update count on first open.
      Office.addin.onVisibilityModeChanged(function (args) {
        if (args.visibilityMode === "Taskpane") {
          updateCount(); // Update count on subsequent opens.
        }
      });
    });

    function updateCount() {
      _count++;
      document.getElementById("run").textContent = "Task pane opened " + _count + " times.";
    }
    ```

1. Enregistrez vos changements et exécutez le projet.

   ```command line
   npm start
   ```

Chaque fois que vous ouvrez le volet Office, le nombre de fois où il a été ouvert est incrémenté. La valeur de **_count** ne sera pas perdue, car le runtime partagé maintient votre code en cours d’exécution même lorsque le volet Office est fermé.

## <a name="runtime-lifetime"></a>Durée de vie de l’exécution

Lorsque vous ajoutez l’élément **\<Runtime\>** , vous spécifiez également une durée de vie avec une valeur ou `long` `short`. Configurez cette valeur sur `long` pour tirer parti de fonctionnalités telles que le démarrage de votre complément lorsque le document s’ouvre, continuer à exécuter un code après la fermeture du volet des tâches, ou utiliser CORS et DOM à partir de fonctions personnalisées.

> [!NOTE]
> La valeur de la durée de vie par défaut est `short`, mais nous vous recommandons d’utiliser `long` dans les compléments Excel, PowerPoint et Word. Si vous avez défini votre runtime sur `short` dans cet exemple, votre complément démarre lorsque vous appuyez sur l’un de vos boutons du ruban, mais il se peut qu’il se ferme une fois l’exécution de votre gestionnaire de ruban terminée. De la même façon, le complément démarre lorsque le volet des tâches est ouvert, mais il se peut se fermer à la fermeture du volet des tâches.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> Si votre complément inclut l’élément **\<Runtimes\>** dans le manifeste (requis pour un runtime partagé) et que les conditions d’utilisation de Microsoft Edge avec WebView2 (basée sur Chromium) sont remplies, il utilise ce contrôle WebView2. Si les conditions ne sont pas remplies, il utilise Internet Explorer 11, quelle que soit la version Windows ou Microsoft 365 version. Pour plus d’informations, consultez [Runtimes](/javascript/api/manifest/runtimes) and [Browsers utilisés par les compléments Office ](../concepts/browsers-used-by-office-web-add-ins.md).

## <a name="about-the-shared-runtime"></a>À propos du runtime partagé

Sur Windows ou Mac, votre complément exécute du code pour les boutons du ruban, les fonctions personnalisées et le volet Office dans des environnements d’exécution distincts. Cela permet de créer des limitations, telles que l'impossibilité de partager aisément des données globales ou de pouvoir accéder à l'ensemble des fonctionnalités CORS à partir d’une fonction personnalisée.

Toutefois, vous pouvez configurer votre complément Office pour partager du code dans le même runtime (également appelé runtime partagé). Vous pouvez ainsi améliorer la coordination dans votre complément et accéder au volet des tâches DOM et CORS à partir de toutes les parties de votre complément.

La configuration d’un runtime partagé permet les scénarios suivants.

- Votre complément Office peut utiliser des fonctionnalités d’interface utilisateur supplémentaires.
  - [Activer et désactiver des commandes de complément](../design/disable-add-in-commands.md)
  - [Exécuter un cote dans votre complément Office lors de l’ouverture du document](run-code-on-document-open.md)
  - [Afficher ou masquer le volet des tâches de votre complément Office](show-hide-add-in.md)
- Les éléments suivants sont disponibles uniquement pour les compléments Excel.
  - [Ajouter des raccourcis clavier personnalisés à votre complément Office (préversion)](../design/keyboard-shortcuts.md)
  - [Créer des onglets contextuels personnalisés dans des compléments Office (préversion)](../design/contextual-tabs.md)
  - Les fonctions personnalisées bénéficieront d'une prise en charge complète de CORS.
  - Les fonctions personnalisées peuvent appeler les API Office.js pour lire les données d’un document feuille de calcul.

Pour Office sur Windows, le runtime partagé utilise Microsoft Edge avec WebView2 (basé sur Chromium) si les conditions de son utilisation sont remplies comme expliqué dans [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md). Sinon, il utilise Internet Explorer 11. De plus, tous les boutons affichés par votre complément sur le ruban s’exécutent dans le même runtime partagé. L’image suivante montre comment les fonctions personnalisées, l’interface utilisateur du ruban et le code du volet Office s’exécutent tous dans le même runtime.

![Diagramme d’une fonction personnalisée, du volet des tâches et des boutons du ruban qui s’exécutent tous dans un runtime de navigateur partagé dans Excel.](../images/custom-functions-in-browser-runtime.png)

### <a name="debug"></a>Débogage

Lors de l’utilisation d’un runtime partagé, vous ne pouvez pas utiliser Visual Studio Code pour déboguer des fonctions personnalisées dans Excel sur Windows à cette date. Vous devez utiliser les outils de développement à la place. Pour plus d’informations, voir [Déboguer des compléments à l’aide des Outils de développement pour Internet Explorer](../testing/debug-add-ins-using-f12-tools-ie.md) ou [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](../testing/debug-add-ins-using-devtools-edge-chromium.md).

### <a name="multiple-task-panes"></a>Multiples volets des tâches

Ne concevez pas votre complément pour utiliser plusieurs volets des tâches si vous envisagez d’utiliser un runtime partagé. Un runtime partagé prend uniquement en charge l’utilisation d’un volet des tâches. Notez que tout volet des tâches sans `<TaskpaneID>` est considéré comme un volet des tâches différent.

## <a name="see-also"></a>Voir aussi

- [Appeler des API Excel à partir d'une fonction personnalisée](../excel/call-excel-apis-from-custom-function.md)
- [Ajouter des raccourcis clavier personnalisés à votre complément Office (préversion)](../design/keyboard-shortcuts.md)
- [Créer des onglets contextuels personnalisés dans des compléments Office (préversion)](../design/contextual-tabs.md)
- [Activer et désactiver des commandes de complément](../design/disable-add-in-commands.md)
- [Exécuter un cote dans votre complément Office lors de l’ouverture du document](run-code-on-document-open.md)
- [Afficher ou masquer le volet des tâches de votre complément Office](show-hide-add-in.md)
- [Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Runtimes dans les compléments Office](../testing/runtimes.md)
