---
title: Personnaliser votre complément compatible avec l’authentification unique Node.js
description: En savoir plus sur la personnalisation du module de personnalisation de LSO que vous avez créé avec le générateur Yeoman.
ms.date: 02/01/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 174df5e58e794b94b02025bd90a65f5ae8e26d44
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234169"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a>Personnaliser votre complément compatible avec l’authentification unique Node.js

> [!IMPORTANT]
> Cet article s’appuie sur le compl?ment sso-enabled qui est créé en compl?tant le démarrage rapide de l' [sign-on unique (SSO).](sso-quickstart.md) Veuillez effectuer le démarrage rapide avant de lire cet article.

Le [](sso-quickstart.md) démarrage rapide de l' cesso crée un add-in ssO qui obtient les informations de profil de l’utilisateur et les écrit dans le document ou le message. Dans cet article, vous allez passer en revue le processus de mise à jour du add-in que vous avez créé avec le générateur Yeoman dans le démarrage rapide de l’eoso, pour ajouter de nouvelles fonctionnalités qui nécessitent différentes autorisations.

## <a name="prerequisites"></a>Configuration requise

- Un add-in Office que vous avez créé en suivant les instructions du démarrage rapide de [l' cesso.](sso-quickstart.md)

- Au moins quelques fichiers et dossiers stockés sur OneDrive Entreprise dans votre abonnement Microsoft 365.

- [Node.js](https://nodejs.org) (la dernière version [LTS](https://nodejs.org/about/releases))

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a>Passer en revue le contenu du projet

Commençons par un examen rapide du projet de add-in que vous avez précédemment créé avec le [générateur Yeoman.](sso-quickstart.md)

> [!NOTE]
> À des endroits où cet article fait référence à des fichiers de script à l’aide de l’extension de fichier **.js,** supposez plutôt l’extension de fichier **.ts** si votre projet a été créé avec TypeScript.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a>Ajouter de nouvelles fonctionnalités

Le add-in que vous avez créé avec le démarrage rapide de l' cesso utilise Microsoft Graph pour obtenir les informations de profil de l’utilisateur et écrit ces informations dans le document ou le message. Nous allons modifier la fonctionnalité du add-in de telle façon qu’il obtient les noms des 10 principaux fichiers et dossiers du OneDrive Entreprise de l’utilisateur et écrit ces informations dans le document ou le message. L’activation de cette nouvelle fonctionnalité nécessite la mise à jour des autorisations d’application dans Azure et la mise à jour du code dans le projet de add-in.

### <a name="update-app-permissions-in-azure"></a>Mettre à jour les autorisations d’application dans Azure

Pour que le module puisse lire correctement le contenu de OneDrive Entreprise de l’utilisateur, ses informations d’inscription d’application dans Azure doivent être mises à jour avec les autorisations appropriées. Pour accorder à l’application **l’autorisation Files.Read.All** et révoquer l’autorisation **User.Read,** qui n’est plus nécessaire, complétez les étapes suivantes.

1. Accédez au [portail Azure et](https://ms.portal.azure.com/#home) **connectez-vous à l’aide** de vos informations d’identification d’administrateur Microsoft 365.

2. Accédez à la page **Inscriptions des applications.**
    > [!TIP]
    > Pour ce faire, vous  pouvez choisir la vignette Inscriptions de l’application sur la page d’accueil Azure ou à l’aide de la zone de recherche de la page d’accueil pour rechercher et choisir les inscriptions **d’applications.**

3. Dans la page **Inscriptions de l’application,** choisissez l’application que vous avez créée lors du démarrage rapide.
    > [!TIP]
    > Le **nom d’affichage** de l’application correspond au nom de la application que vous avez spécifié lors de la création du projet avec le générateur Yeoman.

4. Dans la page vue d’ensemble  de l’application, choisissez les **autorisations d’API** sous le titre Gérer sur le côté gauche de la page.

5. Dans la **ligne User.Read** de la table d’autorisations, choisissez les sélections, puis sélectionnez Révoquer le consentement de l’administrateur dans le menu qui s’affiche. 

6. Sélectionnez **le bouton Oui,** supprimer en réponse à l’invite qui s’affiche.

7. Dans la **ligne User.Read** du tableau des autorisations,  choisissez les sélections, puis sélectionnez Supprimer l’autorisation du menu qui s’affiche.

8. Sélectionnez **le bouton Oui,** supprimer en réponse à l’invite qui s’affiche.

9. Cliquez sur le bouton **Ajouter une autorisation**.

10. Dans le panneau qui s’ouvre, **choisissez Microsoft Graph,** puis les **autorisations déléguées.**

11. Dans le panneau **Demander des autorisations d’API** :

    a. Sous **Fichiers,** **sélectionnez Files.Read.All**.

    b. Sélectionnez **le bouton Ajouter des autorisations** en bas du panneau pour enregistrer ces modifications d’autorisations.

12. Sélectionnez le **bouton Accorder le consentement de l’administrateur pour [nom du client].**

13. Sélectionnez **le bouton** Oui en réponse à l’invite qui s’affiche.

### <a name="update-code-in-the-add-in-project"></a>Mettre à jour le code dans le projet de add-in

Pour permettre au add-in de lire le contenu du OneDrive Entreprise de l’utilisateur, vous devez :

- Mettez à jour le code qui fait référence à l’URL, aux paramètres et à l’étendue d’accès requis de Microsoft Graph.

- Mettez à jour le code qui définit l’interface utilisateur du volet Des tâches, afin qu’il décrive avec précision les nouvelles fonctionnalités.

- Mettez à jour le code qui analyse la réponse de Microsoft Graph et l’écrit dans le document ou le message.

Les étapes suivantes décrivent ces mises à jour.

### <a name="changes-required-for-any-type-of-add-in"></a>Modifications requises pour n’importe quel type de add-in

Pour modifier l’URL, les paramètres et l’étendue d’accès de Microsoft Graph et mettre à jour l’interface utilisateur du volet Des tâches, complétez les étapes suivantes pour votre application. Ces étapes sont les mêmes, quelle que soit l’application Office ciblée par votre application.

1. Dans **le ./. Fichier ENV** :

    a. Remplacez `GRAPH_URL_SEGMENT=/me` par ce qui suit : `GRAPH_URL_SEGMENT=/me/drive/root/children`

    b. Remplacez `QUERY_PARAM_SEGMENT=` par ce qui suit : `QUERY_PARAM_SEGMENT=?$select=name&$top=10`

    c. Remplacez `SCOPE=User.Read` par ce qui suit : `SCOPE=Files.Read.All`

2. Dans **./manifest.xml**, recherchez la ligne vers la fin du fichier et remplacez-la `<Scope>User.Read</Scope>` par la `<Scope>Files.Read.All</Scope>` ligne.

3. Dans **./src/helpers/fallbackauthdialog.js** (ou **dans ./src/helpers/fallbackauthdialog.ts** pour un projet TypeScript), recherchez la chaîne et remplacez-la par la chaîne définie comme suit `https://graph.microsoft.com/User.Read` `https://graph.microsoft.com/Files.Read.All` `requestObj` :

    ```javascript
    var requestObj = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

    ```typescript
    var requestObj: Object = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

4. Dans **./src/taskpane/taskpane.html**, recherchez l’élément et mettez à jour le texte dans cet élément pour décrire les nouvelles fonctionnalités `<section class="ms-firstrun-instructionstep__header">` du module.

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. Dans **./src/taskpane/taskpane.html**, recherchez et remplacez les deux occurrences de la chaîne `Get My User Profile Information` par la chaîne `Read my OneDrive for Business` .

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">Click the <b>Read my OneDrive for Business</b>
            button.</span>
        <div class="clearfix"></div>
    </li>
    ```

    ```html
    <p align="center">
        <button id="getGraphDataButton" class="popupButton ms-Button ms-Button--primary"><span
                class="ms-Button-label">Read my OneDrive for Business</span></button>
    </p>
    ```

6. Dans **./src/taskpane/taskpane.html**, recherchez et remplacez la chaîne `Your user profile information will be displayed in the document.` par la chaîne `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.` .

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. Mettez à jour le code qui analyse la réponse de Microsoft Graph et l’écrit dans le document ou le message, en suivant les instructions de la section qui correspond à votre type de add-in :

    - [Modifications requises pour un add-in Excel (JavaScript)](#changes-required-for-an-excel-add-in-javascript)
    - [Modifications requises pour un add-in Excel (TypeScript)](#changes-required-for-an-excel-add-in-typescript)
    - [Modifications requises pour un add-in Outlook (JavaScript)](#changes-required-for-an-outlook-add-in-javascript)
    - [Modifications requises pour un add-in Outlook (TypeScript)](#changes-required-for-an-outlook-add-in-typescript)
    - [Modifications requises pour un add-in PowerPoint (JavaScript)](#changes-required-for-a-powerpoint-add-in-javascript)
    - [Modifications requises pour un add-in PowerPoint (TypeScript)](#changes-required-for-a-powerpoint-add-in-typescript)
    - [Modifications requises pour un add-in Word (JavaScript)](#changes-required-for-a-word-add-in-javascript)
    - [Modifications requises pour un add-in Word (TypeScript)](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a>Modifications requises pour un add-in Excel (JavaScript)

Si votre add-in est un add-in Excel créé avec JavaScript, a apporté les modifications suivantes dans **./src/helpers/documentHelper.js**:

1. Recherchez `writeDataToOfficeDocument` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToExcel(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. Recherchez `filterUserProfileInfo` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. Recherchez `writeDataToExcel` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    function writeDataToExcel(result) {
      return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            let innerArray = [];
            innerArray.push(oneDriveInfo[i]);
            data.push(innerArray);
          }
        }

        const rangeAddress = `B5:B${5 + (data.length - 1)}`;
        const range = sheet.getRange(rangeAddress);
        range.values = data;
        range.format.autofitColumns();

        return context.sync();
      });
    }
    ```

4. Supprimez la `writeDataToOutlook` fonction.

5. Supprimez la `writeDataToPowerPoint` fonction.

6. Supprimez la `writeDataToWord` fonction.

Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.

### <a name="changes-required-for-an-excel-add-in-typescript"></a>Modifications requises pour un add-in Excel (TypeScript)

Si votre add-in est un module excel créé avec TypeScript, ouvrez **./src/taskpane/taskpane.ts,** recherchez la fonction et remplacez-la par la fonction suivante `writeDataToOfficeDocument` :

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    }

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        let innerArray = [];
        innerArray.push(itemNames[i]);
        data.push(innerArray);
      }
    }

    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.

### <a name="changes-required-for-an-outlook-add-in-javascript"></a>Modifications requises pour un add-in Outlook (JavaScript)

Si votre add-in est un add-in Outlook créé avec JavaScript, a apporté les modifications suivantes dans **./src/helpers/documentHelper.js**:

1. Recherchez `writeDataToOfficeDocument` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToOutlook(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to message. " + error.toString()));
        }
      });
    }
    ```

2. Recherchez `filterUserProfileInfo` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. Recherchez `writeDataToOutlook` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    function writeDataToOutlook(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
      }

      Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
    }
    ```

4. Supprimez la `writeDataToExcel` fonction.

5. Supprimez la `writeDataToPowerPoint` fonction.

6. Supprimez la `writeDataToWord` fonction.

Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.

### <a name="changes-required-for-an-outlook-add-in-typescript"></a>Modifications requises pour un add-in Outlook (TypeScript)

Si votre add-in est un add-in Outlook qui a été créé avec TypeScript, ouvrez **./src/taskpane/taskpane.ts**, recherchez la fonction et remplacez-la par la fonction suivante `writeDataToOfficeDocument` :

```typescript
export function writeDataToOfficeDocument(result: Object): void {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
        if (itemNames[i] !== null) {
        data.push(itemNames[i]);
        }
    }

    let objectNames: string = "";
    for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
    }

    Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
}
```

Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a>Modifications requises pour un add-in PowerPoint (JavaScript)

Si votre add-in est un add-in PowerPoint créé avec JavaScript, a apporté les modifications suivantes dans **./src/helpers/documentHelper.js**:

1. Recherchez `writeDataToOfficeDocument` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToPowerPoint(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. Recherchez `filterUserProfileInfo` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. Recherchez `writeDataToPowerPoint` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    function writeDataToPowerPoint(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "\n";
      }

      Office.context.document.setSelectedDataAsync(
        objectNames, 
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            throw asyncResult.error.message;
          }
      });
    }
    ```

4. Supprimez la `writeDataToExcel` fonction.

5. Supprimez la `writeDataToOutlook` fonction.

6. Supprimez la `writeDataToWord` fonction.

Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a>Modifications requises pour un add-in PowerPoint (TypeScript)

Si votre add-in est un add-in PowerPoint qui a été créé avec TypeScript, ouvrez **./src/taskpane/taskpane.ts**, recherchez la fonction et remplacez-la par la fonction suivante `writeDataToOfficeDocument` :

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];

  let itemNames: string[] = [];
  let oneDriveItems = result["value"];
  for (let item of oneDriveItems) {
    itemNames.push(item["name"]);
  };

  for (let i = 0; i < itemNames.length; i++) {
    if (itemNames[i] !== null) {
      data.push(itemNames[i]);
    }
  }

  let objectNames: string = "";
  for (let i = 0; i < data.length; i++) {
    objectNames += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(objectNames, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.

### <a name="changes-required-for-a-word-add-in-javascript"></a>Modifications requises pour un add-in Word (JavaScript)

Si votre add-in est un add-in Word créé avec JavaScript, a apporté les modifications suivantes dans **./src/helpers/documentHelper.js**:

1. Recherchez `writeDataToOfficeDocument` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToWord(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. Recherchez `filterUserProfileInfo` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. Recherchez `writeDataToWord` la fonction et remplacez-la par la fonction suivante :

    ```javascript
    function writeDataToWord(result) {
      return Word.run(function (context) {
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            data.push(oneDriveInfo[i]);
          }
        }

        const documentBody = context.document.body;
        for (let i = 0; i < data.length; i++) {
          if (data[i] !== null) {
            documentBody.insertParagraph(data[i], "End");
          }
        }

        return context.sync();
      });
    }
    ```

4. Supprimez la `writeDataToExcel` fonction.

5. Supprimez la `writeDataToOutlook` fonction.

6. Supprimez la `writeDataToPowerPoint` fonction.

Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.

### <a name="changes-required-for-a-word-add-in-typescript"></a>Modifications requises pour un add-in Word (TypeScript)

Si votre add-in est un add-in Word créé avec TypeScript, ouvrez **./src/taskpane/taskpane.ts,** recherchez la fonction et remplacez-la par la fonction suivante `writeDataToOfficeDocument` :

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function(context) {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        data.push(itemNames[i]);
      }
    }

    const documentBody: Word.Body = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
```

Une fois ces modifications apportées, continuez à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.

## <a name="try-it-out"></a>Try it out

Si votre compl?ment est un compl?ment Excel, Word ou PowerPoint, compl?ez les étapes de la section suivante pour l’essayer. Si votre compl?ment est un compl?ment Outlook, compl?ez les étapes dans la section [Outlook.](#outlook)

### <a name="excel-word-and-powerpoint"></a>Excel, Word et PowerPoint

Pour tester un complément Excel, Word ou PowerPoint, procédez comme suit.

1. Dans le dossier racine du projet, exécutez la commande suivante pour créer le projet, démarrez le serveur web local et chargez une version test de votre application dans l’application cliente Office précédemment sélectionnée.

    > [!NOTE]
    > Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.

    ```command&nbsp;line
    npm start
    ```

2. Dans l’application cliente Office qui s’ouvre lorsque vous exécutez la commande précédente (c’est-à-dire, Excel, Word ou PowerPoint), assurez-vous que vous êtes connecté avec un utilisateur membre [](sso-quickstart.md#configure-sso) de la même organisation Microsoft 365 que le compte d’administrateur Microsoft 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’ouvrez-vous pour l’application. Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique. 

3. Dans l’application client Office, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément. L’image ci-après illustre ce bouton dans Excel.

    ![Screenshot showing highlighted add-in button in Excel ribbon](../images/excel-quickstart-addin-3b.png)

4. En bas du volet Des tâches, sélectionnez le bouton Lire **mon OneDrive** Entreprise pour lancer le processus d’pertinence.

5. Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs. Cela peut se produire lorsque l’administrateur client n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Microsoft 365 (professionnel ou scolaire). Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.

    ![Capture d’écran montrant la boîte de dialogue des autorisations demandées avec le bouton Accepter mis en évidence](../images/sso-permissions-request.png)

    > [!NOTE]
    > Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.

6. Le add-in lit les données du OneDrive Entreprise de l’utilisateur et écrit les noms des 10 principaux fichiers et dossiers dans le document. L’image suivante montre un exemple de noms de fichiers et de dossiers écrits dans une feuille de calcul Excel.

    ![Capture d’écran montrant les informations OneDrive Entreprise dans la feuille de calcul Excel](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a>Outlook

Pour tester un complément Outlook, procédez comme suit.

1. Dans le dossier racine du projet, exécutez la commande suivante pour créer le projet, démarrez le serveur web local et chargez une version test de votre application. 

    > [!NOTE]
    > Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman. Il se peut également que vous deviez exécuter votre invite de commande ou votre terminal en tant qu'administrateur pour que les modifications soient effectuées.

    ```command&nbsp;line
    npm start
    ```

2. Assurez-vous que vous êtes connecté à Outlook avec un utilisateur membre de la même organisation Microsoft 365 que le compte d’administrateur Microsoft 365 que celui que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’oD [SSO](sso-quickstart.md#configure-sso) pour l’application. Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.

3. Rédigez un nouveau message dans Outlook.

4. Dans la fenêtre de composition du message, choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet du complément.

    ![Capture d’écran illustrant la fenêtre Outlook Composer un message et le bouton du ruban du complément mis en évidence](../images/outlook-sso-ribbon-button.png)

5. En bas du volet Des tâches, sélectionnez le bouton Lire **mon OneDrive** Entreprise pour lancer le processus d’pertinence.

6. Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs. Cela peut se produire lorsque l’administrateur client n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Microsoft 365 (professionnel ou scolaire). Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.

    ![Capture d’écran de la boîte de dialogue des autorisations demandées avec le bouton Accepter mis en évidence](../images/sso-permissions-request.png)

    > [!NOTE]
    > Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.

7. Le add-in lit les données du OneDrive Entreprise de l’utilisateur et écrit les noms des 10 principaux fichiers et dossiers dans le corps du message électronique.

    ![Capture d’écran montrant les informations OneDrive Entreprise dans la fenêtre composer un message Outlook](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez personnalisé avec succès la fonctionnalité du module de personnalisation de l’oDS que vous avez créée avec le générateur Yeoman dans le démarrage rapide de l’personnalisation [SSO.](sso-quickstart.md) Pour en savoir plus sur les étapes de configuration de l’authentification unique effectuées automatiquement par le générateur Yeoman et le code facilitant le processus d’authentification unique, veuillez consultez le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="see-also"></a>Consultez aussi

- [Activer l’authentification unique pour des compléments Office](../develop/sso-in-office-add-ins.md)
- [Démarrage rapide de l’authentification unique (SSO)](sso-quickstart.md)
- [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md)
- [Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)
