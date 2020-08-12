---
title: 'Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office'
description: Découvrez comment partager des données et des événements entre des fonctions personnalisées et le volet Office dans Excel.
ms.date: 05/17/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: a48d43270787648d8e5a53c885eab4b69cd8842e
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641150"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a>Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office

Vous pouvez configurer votre complément Excel pour utiliser un runtime partagé. Vous pouvez ainsi partager des données globales ou envoyer des événements entre le volet des tâches et les fonctions personnalisées.

Pour la plupart des scénarios de fonctions personnalisées, nous vous recommandons d’utiliser un runtime partagé, sauf si vous avez une raison particulière d’utiliser une fonction personnalisée (sans interface utilisateur).

Ce didacticiel part du principe que vous avez l’habitude d’utiliser le générateur d’Office YÔ pour créer des projets de complément. Envisagez d’effectuer le [Didacticiel sur les fonctions Excel personnalisées](./excel-tutorial-create-custom-functions.md), si ce n’est déjà fait.

## <a name="create-the-add-in-project"></a>Création du projet de complément

Utilisez le générateur Yeoman pour créer un projet de complément Excel. Exécutez la commande suivante, puis répondez aux invites avec les réponses suivantes :

```command line
yo office
```

- Choose a project type (Choisissez un type de projet) : **projet de complément Fonctions personnalisées Excel**
- Choose a script type (Choisissez un type de script) :  **JavaScript**
- What do you want to name your add-in? (Comment souhaitez-vous nommer votre complément ?)  **My Office Add-in**

![Capture d’écran de réponse aux invites à partir d’Office pour créer le projet de complément.](../images/yo-office-excel-project.png)

Après avoir exécuté l’Assistant, le générateur crée le projet et installe les composants Node de prise en charge.

## <a name="configure-the-manifest"></a>Configurer le manifeste

1. Démarrez Visual Studio Code et ouvrez le projet **My Office Add-in**.
2. Ouvrez le fichier **manifest.xml**.
3. Recherchez la section `<VersionOverrides>`, puis ajoutez l'exemple d'entrée suivante à la section `<Runtimes>`. La durée de vie doit être **longue** afin que les fonctions personnalisées puissent continuer de fonctionner même quand le volet Office est fermé.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

4. Dans l’élément `<Page>`, remplacez l’emplacement de la source **Functions.Page.Url** par **ContosoAddin.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. Dans la section `<DesktopFormFactor>`, changez la valeur **Commands.Url** de **FunctionFile** pour utiliser **ContosoAddin.Url**.

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. Dans la section `<Action>`, remplacez l’emplacement de la source **Taskpane.Url** par **ContosoAddin.Url**.

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. Ajoutez un nouvel **ID d’URL** pour **ContosoAddin.Url** qui pointe vers **taskpane.html**.

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. Enregistrez vos changements et regénérez le projet.

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a>Partager l’état entre une fonction personnalisée et du code du volet Office

À présent que les fonctions personnalisées s’exécutent dans le même contexte que votre code du volet Office, elles peuvent partager l’état directement, sans utiliser l’objet **Storage**. Les instructions suivantes montrent comment partager une variable globale entre une fonction personnalisée et du code du volet Office.

### <a name="create-custom-functions-to-get-or-store-shared-state"></a>Créer des fonctions personnalisées pour obtenir ou stocker l’état partagé

1. Dans Visual Studio Code, ouvrez le fichier **src/functions/functions.js**.
2. Sur la ligne 1, tout en haut, insérez le code suivant. Cette opération initialise une variable globale nommée **sharedState**.

   ```js
   window.sharedState = "empty";
   ```

3. Ajoutez le code suivant pour créer une fonction personnalisée qui stocke des valeurs dans la variable **sharedState**.

   ```js
   /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
   function storeValue(sharedValue) {
     window.sharedState = sharedValue;
     return "value stored";
   }
   ```

4. Ajoutez le code suivant pour créer une fonction personnalisée qui obtient la valeur actuelle de la variable **sharedState**.

   ```js
   /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
   function getValue() {
     return window.sharedState;
   }
   ```

5. Enregistrez le fichier.

### <a name="create-task-pane-controls-to-work-with-global-data"></a>Créer des contrôles du volet Office pour utiliser des données globales

1. Ouvrez le fichier **src/taskpane/taskpane.html**.
2. Ajoutez l'élément de script suivant juste avant l’élément `</head>`.

   ```html
   <script src="functions.js"></script>
   ```

3. Après l’élément de fermeture `</main>`, ajoutez le code HTML suivant. Le code HTML crée deux zones de texte et des boutons permettant d’obtenir ou de stocker des données globales.

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
     <li>Select <strong>Get</strong> to display the value in the task pane.</li>
   </ol>
   <p>Store new value to shared state</p>
   <div>
     <input type="text" id="storeBox" />
     <button onclick="storeSharedValue()">Store</button>
   </div>

   <p>Get shared state value</p>
   <div>
     <input type="text" id="getBox" />
     <button onclick="getSharedValue()">Get</button>
   </div>
   ```

4. Avant l’élément `<body>`, ajoutez le code suivant. Ce code gère les événements de clic de bouton quand l’utilisateur souhaite stocker ou obtenir des données globales.

   ```js
   <script>
   function storeSharedValue() {
   let sharedValue = document.getElementById('storeBox').value;
   window.sharedState = sharedValue;
   }

   function getSharedValue() {
   document.getElementById('getBox').value = window.sharedState;
   }</script>
   ```

5. Enregistrez le fichier.
6. Générez le projet.

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a>Essayer de partager des données entre les fonctions personnalisées et le volet Office

- Démarrez le projet à l’aide de la commande suivante.

  ```command line
  npm run start
  ```

Une fois Excel démarré, vous pouvez utiliser les boutons du volet Office pour stocker ou obtenir des données partagées. Entrez `=CONTOSO.GETVALUE()` dans une cellule pour que la fonction personnalisée extraie les mêmes données partagées. Vous pouvez également utiliser `=CONTOSO.STOREVALUE("new value")` pour remplacer les données partagées par une nouvelle valeur.

> [!NOTE]
> La configuration de votre projet comme illustré dans cet article permet de partager le contexte entre des fonctions personnalisées et le volet Office. L’appel de certaines API Office à partir de fonctions personnalisées est possible. Pour plus d’informations, [consultez Appeler les API Microsoft Excel à partir d’une fonction personnalisée](../excel/call-excel-apis-from-custom-function.md).
