---
title: Ajouter la fonctionnalité Microsoft Graph à votre projet de démarrage rapide de l’authentification unique
description: Découvrez comment ajouter de nouvelles fonctionnalités Microsoft Graph au complément prenant en charge l’authentification unique que vous avez créé.
ms.date: 06/10/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 6f8784dae3f947baaedc3232e06a5208988ba9e9
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091138"
---
# <a name="add-microsoft-graph-functionality-to-your-sso-quick-start-project"></a>Ajouter la fonctionnalité Microsoft Graph à votre projet de démarrage rapide de l’authentification unique

> [!IMPORTANT]
> Cet article s’appuie sur le complément prenant en charge l’authentification unique créé en effectuant le démarrage rapide de l’authentification [unique (SSO](sso-quickstart.md)). Veuillez suivre le guide de démarrage rapide avant de lire cet article.

Le [démarrage rapide de](sso-quickstart.md) l’authentification unique crée un complément prenant en charge l’authentification unique qui obtient les informations de profil de l’utilisateur connecté et les écrit dans le document ou le message. Dans cet article, vous allez parcourir le processus de mise à jour du complément que vous avez créé avec le générateur Yeoman dans le démarrage rapide de l’authentification unique, pour ajouter de nouvelles fonctionnalités qui nécessitent différentes autorisations.

## <a name="prerequisites"></a>Conditions préalables

- Complément Office que vous avez créé en suivant les instructions du [guide de démarrage rapide de l’authentification unique](sso-quickstart.md).

- Au moins quelques fichiers et dossiers stockés sur OneDrive Entreprise dans votre abonnement Microsoft 365.

- [Node.js](https://nodejs.org) (la dernière version [LTS](https://nodejs.org/about/releases))

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a>Passer en revue le contenu du projet

Commençons par un rapide examen du projet de complément que vous avez créé précédemment [avec le générateur Yeoman](sso-quickstart.md).

> [!NOTE]
> Dans les endroits où cet article référence des fichiers de script à l’aide **.js'extension** de fichier, supposons plutôt l’extension **de fichier .ts** si votre projet a été créé avec TypeScript.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a>Ajouter de nouvelles fonctionnalités

Le complément que vous avez créé avec le démarrage rapide de l’authentification unique utilise Microsoft Graph pour obtenir les informations de profil de l’utilisateur connecté et les écrit dans le document ou le message. Modifions les fonctionnalités du complément de sorte qu’il obtient les noms des 10 fichiers et dossiers les plus utilisés à partir du OneDrive Entreprise de l’utilisateur connecté et qu’il écrit ces informations dans le document ou le message. L’activation de cette nouvelle fonctionnalité nécessite la mise à jour des autorisations d’application dans Azure et la mise à jour du code dans le projet de complément.

### <a name="update-app-permissions-in-azure"></a>Mettre à jour les autorisations d’application dans Azure

Pour que le complément puisse lire le contenu du OneDrive Entreprise de l’utilisateur, ses informations d’inscription d’application dans Azure doivent être mises à jour avec les autorisations appropriées. Effectuez les étapes suivantes pour accorder à l’application l’autorisation **Files.Read.All** et révoquer l’autorisation **User.Read** , qui n’est plus nécessaire.

1. Connectez-vous au [Portail Azure](https://portal.azure.com) avec vos **informations d’identification d’administrateur Microsoft 365**.

1. Accédez à la page **inscriptions d'applications** et choisissez l’inscription d’application que vous avez créée au démarrage rapide.
    > [!TIP]
    > Le **nom d’affichage** de l’application correspond au nom de complément que vous avez spécifié lors de la création du projet avec le générateur Yeoman.

1. Sous **Gérer**, choisissez **autorisations d’API**.

1. Dans la ligne **User.Read** de la table d’autorisations, choisissez les points de suspension, puis sélectionnez **Révoquer le consentement administrateur** dans le menu qui s’affiche.

    :::image type="content" source="../images/app-registration-revoke-admin-consent.png" alt-text="Capture d’écran du bouton Révoquer le consentement de l’administrateur sur la page d’autorisations de l’API.":::

1. Sélectionnez le bouton **Oui, supprimer** en réponse à l’invite affichée.

1. Dans la ligne **User.Read** de la table d’autorisations, choisissez les points de suspension, puis **sélectionnez Supprimer l’autorisation** dans le menu qui s’affiche.

    :::image type="content" source="../images/app-registration-remove-permission.png" alt-text="Capture d’écran du bouton Supprimer l’autorisation dans la page Autorisations de l’API.":::

1. Sélectionnez le bouton **Oui, supprimer** en réponse à l’invite affichée.

1. Cliquez sur le bouton **Ajouter une autorisation**.

1. Dans le panneau qui s’ouvre, choisissez **Microsoft Graph**, puis choisissez **Autorisations déléguées**.

1. Dans le panneau **Demander des autorisations d’API** :

    a. Sous **Fichiers**, sélectionnez **Files.Read.All**.

    b. Sélectionnez le bouton **Ajouter des autorisations** en bas du panneau pour enregistrer ces modifications d’autorisations.

1. Sélectionnez le bouton **Accorder le consentement administrateur pour [nom du locataire** ].

1. Sélectionnez le bouton **Oui** en réponse à l’invite affichée.

### <a name="update-code-in-the-add-in-project"></a>Mettre à jour le code dans le projet de complément

Pour permettre au complément de lire le contenu du OneDrive Entreprise de l’utilisateur connecté, vous devez :

- Mettez à jour le code qui fait référence à l’URL, aux paramètres et à l’étendue d’accès requis de Microsoft Graph.

- Mettez à jour le code qui définit l’interface utilisateur du volet Office, afin qu’il décrive avec précision les nouvelles fonctionnalités.

- Mettez à jour le code qui analyse la réponse de Microsoft Graph et l’écrit dans le document ou le message.

Les étapes suivantes décrivent ces mises à jour.

### <a name="changes-required-for-any-type-of-add-in"></a>Modifications requises pour n’importe quel type de complément

Effectuez les étapes suivantes pour votre complément, afin de modifier l’URL, les paramètres et l’étendue d’accès microsoft Graph et de mettre à jour l’interface utilisateur du volet Office. Ces étapes sont les mêmes, quelle que soit l’application Office vos cibles de complément.

1. Dans le **./. Fichier ENV** :

    a. Remplacer `GRAPH_URL_SEGMENT=/me` par `GRAPH_URL_SEGMENT=/me/drive/root/children`

    b. Remplacer `QUERY_PARAM_SEGMENT=` par `QUERY_PARAM_SEGMENT=?$select=name&$top=10`

    c. Remplacer `SCOPE=User.Read` par `SCOPE=Files.Read.All`

1. Dans **./manifest.xml**, recherchez la ligne `<Scope>User.Read</Scope>` près de la fin du fichier et remplacez-la par la ligne `<Scope>Files.Read.All</Scope>`.

1. Dans **./src/helpers/fallbackauthdialog.js** (ou dans **./src/helpers/fallbackauthdialog.ts** pour un projet TypeScript), recherchez la chaîne `https://graph.microsoft.com/User.Read` et remplacez-la par la chaîne `https://graph.microsoft.com/Files.Read.All`, telle que `requestObj` définie comme suit :

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

1. Dans **./src/taskpane/taskpane.html**, recherchez l’élément `<section class="ms-firstrun-instructionstep__header">` et mettez à jour le texte dans cet élément pour décrire les nouvelles fonctionnalités du complément.

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

1. Dans **./src/taskpane/taskpane.html**, recherchez les deux occurrences de la chaîne `Get My User Profile Information` et remplacez-la `Read my OneDrive for Business`par .

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

1. Dans **./src/taskpane/taskpane.html**, recherchez la chaîne `Your user profile information will be displayed in the document.` et remplacez-la `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`par .

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

1. Mettez à jour le code qui analyse la réponse de Microsoft Graph et l’écrit dans le document ou le message, en suivant les instructions de la section qui correspondent à votre type de complément :

    - [Modifications requises pour un complément Office (JavaScript)](#changes-required-for-an-office-add-in-javascript)
    - [Modifications requises pour un complément Office (TypeScript)](#changes-required-for-an-office-add-in-typescript)

### <a name="changes-required-for-an-office-add-in-javascript"></a>Modifications requises pour un complément Office (JavaScript)

Si votre complément Office généré utilise JavaScript, apportez les modifications suivantes dans **./src/helpers/documentHelper.js**.

1. Recherchez la `filterUserProfileInfo` fonction et remplacez-la par la fonction suivante.

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

1. Recherchez `filterUserProfileInfo` et remplacez-le `filterOneDriveInfo`par . Quatre instances doivent être remplacées.

1. Enregistrez les modifications.

Une fois que vous avez apporté ces modifications, passez directement à la section [Essayer](#try-it-out) de cet article pour essayer votre complément mis à jour.

### <a name="changes-required-for-an-office-add-in-typescript"></a>Modifications requises pour un complément Office (TypeScript)

Si votre complément Office généré utilise TypeScript, ouvrez **./src/taskpane/taskpane.ts**.

1. Recherchez la `writeDataToOfficeDocument` fonction et remplacez-la par le code suivant en fonction de l’hôte Office votre complément (Excel, Outlook, Word ou PowerPoint)

#### <a name="excel-code"></a>code Excel

```typescript
  export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[][];

    // Get just the filenames from results
    data = result["value"].map((item) => {
      return [item.name];
    });

    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

#### <a name="outlook-code"></a>code Outlook

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  // Get just the filenames from results.
  const data: string[] = result["value"].map((item) => {
    return item.name;
  });

  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "</br>";
  }
  Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
}
```

#### <a name="word-code"></a>Code Word

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function (context) {
    // Get just the filenames from results.
    const data: string[] = result["value"].map((item) => {
      return item.name;
    });

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

#### <a name="powerpoint-code"></a>code PowerPoint

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  // Get just the filenames from results.
  const data: string[] = result["value"].map((item) => {
    return item.name;
  });
  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(userInfo, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

## <a name="try-it-out"></a>Essayez

Si votre complément est un complément Excel, Word ou PowerPoint, effectuez les étapes décrites dans la section suivante pour l’essayer. Si votre complément est un complément Outlook, effectuez plutôt les étapes de la section [Outlook](#outlook).

### <a name="excel-word-and-powerpoint"></a>Excel, Word et PowerPoint

Pour tester un complément Excel, Word ou PowerPoint, procédez comme suit.

1. Dans le dossier racine du projet, exécutez la commande suivante pour générer le projet, démarrer le serveur web local et charger votre complément dans l’application cliente Office précédemment sélectionnée.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. Dans l’application cliente Office qui s’ouvre lorsque vous exécutez la commande précédente (par exemple, Excel, Word ou PowerPoint), assurez-vous que vous êtes connecté avec un utilisateur membre de la même organisation Microsoft 365 que le compte d’administrateur Microsoft 365 que vous avez utilisé pour vous connecter à Azure lors de la [configuration](sso-quickstart.md#configure-sso) de l’authentification unique  pour l’application. Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique. 

3. Dans l’application cliente Office, choisissez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet des tâches** dans le ruban pour ouvrir le volet des tâches du complément. L’image suivante montre ce bouton dans Excel.

    ![Capture d’écran montrant le bouton de complément mis en surbrillance dans Excel ruban.](../images/excel-quickstart-addin-3b.png)

4. En bas du volet Office, choisissez le bouton **Lire mon OneDrive Entreprise** pour lancer le processus d’authentification unique.

5. Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs. Cela peut se produire lorsque l’administrateur client n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Microsoft 365 (professionnel ou scolaire). Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.

    ![Capture d’écran montrant la boîte de dialogue des autorisations demandées avec le bouton Accepter mis en évidence.](../images/sso-permissions-request.png)

    > [!NOTE]
    > Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.

6. Le complément lit les données du OneDrive Entreprise de l’utilisateur connecté et écrit les noms des 10 principaux fichiers et dossiers dans le document. L’image suivante montre un exemple de noms de fichiers et de dossiers écrits dans une feuille de calcul Excel.

    ![Capture d’écran montrant OneDrive Entreprise informations dans Excel feuille de calcul.](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a>Outlook

Pour tester un complément Outlook, procédez comme suit.

1. Dans le dossier racine du projet, exécutez la commande suivante pour générer le projet, démarrer le serveur web local et charger de manière indépendante votre complément. 

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. Vérifiez que vous êtes connecté à Outlook avec un utilisateur membre de la même organisation Microsoft 365 que le compte d’administrateur Microsoft 365 que vous avez utilisé pour vous connecter à Azure lors [de la configuration de l’authentification](sso-quickstart.md#configure-sso) unique pour l’application. Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.

3. Rédigez un nouveau message dans Outlook.

4. Dans la fenêtre de composition du message, choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet du complément.

    ![Capture d’écran illustrant la fenêtre Outlook Composer un message et le bouton du ruban du complément mis en évidence.](../images/outlook-sso-ribbon-button.png)

5. En bas du volet Office, choisissez le bouton **Lire mon OneDrive Entreprise** pour lancer le processus d’authentification unique.

6. Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs. Cela peut se produire lorsque l’administrateur client n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Microsoft 365 (professionnel ou scolaire). Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.

    ![Capture d’écran de la boîte de dialogue des autorisations demandées avec le bouton Accepter mis en évidence.](../images/sso-permissions-request.png)

    > [!NOTE]
    > Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.

7. Le complément lit les données du OneDrive Entreprise de l’utilisateur connecté et écrit les noms des 10 principaux fichiers et dossiers dans le corps du message électronique.

    ![Capture d’écran montrant OneDrive Entreprise informations dans Outlook fenêtre de composition de message.](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a>Prochaines étapes

Félicitations, vous avez correctement personnalisé les fonctionnalités du complément prenant en charge l’authentification unique que vous avez créé avec le générateur Yeoman dans le guide de [démarrage rapide](sso-quickstart.md) de l’authentification unique. Pour en savoir plus sur les étapes de configuration de l’authentification unique effectuées automatiquement par le générateur Yeoman et le code facilitant le processus d’authentification unique, veuillez consultez le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="see-also"></a>Consultez aussi

- [Activer l’authentification unique pour des compléments Office](../develop/sso-in-office-add-ins.md)
- [Démarrage rapide de l’authentification unique (SSO)](sso-quickstart.md)
- [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md)
- [Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)
- [Utilisation de Visual Studio Code pour publier](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)