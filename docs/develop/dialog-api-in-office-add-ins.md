---
title: Utiliser l’API de boîte de dialogue Office dans vos compléments Office
description: Découvrez les principes de base de la création d’une boîte de dialogue dans un complément Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4dc1bc0b45bb41952cd2ab83fcd62633d598ab4e
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810014"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Utiliser l’API de boîte de dialogue Office dans les compléments Office

Vous pouvez utiliser l’[API de dialogue Office](/javascript/api/office/office.ui) pour ouvrir des boîtes de dialogue dans votre complément Office. Cet article fournit des conseils concernant l’utilisation de l’API de dialogue dans votre complément Office.

> [!NOTE]
> Pour plus d’informations sur les compléments où l’API de dialogue est actuellement prise en charge, consultez la rubrique relative aux [ensembles de conditions requises de l’API de dialogue](/javascript/api/requirement-sets/common/dialog-api-requirement-sets). L’API boîte de dialogue est actuellement prise en charge pour Excel, PowerPoint et Word. La prise en charge d’Outlook est incluse dans différents ensembles de conditions requises&mdash;de boîte aux lettres, consultez la référence de l’API pour plus d’informations.

Un scénario principal pour l’API de dialogue consiste à activer l’authentification à l'aide d'une ressource telle que Google, Facebook, ou Microsoft Graph. Pour plus d’informations, voir [S’authentifier auprès de l'API de boîte de dialogue Office](auth-with-office-dialog-api.md) *une fois* que vous êtes familiarisé(e) avec cet article.

Envisagez d’ouvrir une boîte de dialogue à partir d’un volet Office, d’un complément de contenu ou d’un [complément de commande](../design/add-in-commands.md) pour effectuer les opérations suivantes :

- Afficher les pages de connexion qui ne peuvent pas être ouvertes directement dans un volet Office.
- fournir davantage d’espace à l’écran, ou même un plein écran, pour certaines tâches exécutées dans votre complément ;
- héberger une vidéo qui serait trop petite si elle était limitée à un volet Office.

> [!NOTE]
> Comme des éléments d’interface utilisateur qui se chevauchent peuvent gêner des utilisateurs, évitez d’ouvrir une boîte de dialogue à partir d’un volet Office à moins que votre scénario l’exige. Lorsque vous envisagez d’utiliser la surface d’exposition d’un volet Office, tenez compte du fait que les volets Office peuvent être affichés sous forme d’onglets. Pour obtenir un exemple de volet Office à onglets, consultez l’exemple [De complément Excel JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) .

L’image suivante montre un exemple de boîte de dialogue.

![Boîte de dialogue avec 3 options de connexion affichées devant Word.](../images/auth-o-dialog-open.png)

Notez que la boîte de dialogue s’ouvre toujours au centre de l’écran. L’utilisateur peut la déplacer et la redimensionner. La fenêtre n’est *pas modifiée* : un utilisateur peut continuer à interagir avec le document dans l’application Office et avec la page dans le volet Office, le cas échéant.

## <a name="open-a-dialog-box-from-a-host-page"></a>Ouvrir une boîte de dialogue à partir d’une page hôte

Les API JavaScript Office incluent un objet [Dialog](/javascript/api/office/office.dialog) et deux fonctions dans l’[espace de noms Office.context.ui](/javascript/api/office/office.ui).

Pour ouvrir une boîte de dialogue, généralement une page dans un volet des tâches, votre code appelle la méthode [displayDialogAsync](/javascript/api/office/office.ui) et lui transmet l’URL de la ressource que vous voulez ouvrir. La page sur laquelle cette méthode est appelée est connue sous le nom de « page hôte ». Par exemple, si vous appelez cette méthode dans le script sur index.html d'un volet de tâches, la page index.html correspond à la page hôte de la boîte de dialogue ouverte par la méthode.

La ressource ouverte dans la boîte de dialogue correspond généralement à une page, mais ce peut être une méthode du contrôleur dans une application MVC, un itinéraire, une méthode de service web ou toute autre ressource. Dans cet article, les termes « page » ou « site web » font référence à la ressource dans la boîte de dialogue. Le code suivant est un exemple simple.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
>
> - L’URL utilise le protocole HTTP **S**. Ceci est obligatoire pour toutes les pages chargées dans une boîte de dialogue, pas seulement la première page chargée.
> - Le domaine de la boîte de dialogue est le même que celui de la page hôte, qui peut être la page d’un volet Office ou le [fichier de fonctions](/javascript/api/manifest/functionfile) d’une commande de complément. Obligatoire : la page, la méthode du contrôleur ou toute autre ressource qui est transmise à la méthode `displayDialogAsync` doit se trouver dans le même domaine que la page hôte.

> [!IMPORTANT]
> La page hôte et les ressources s'ouvrant dans la boîte de dialogue doivent avoir le même domaine complet. Si vous tentez de transmettre `displayDialogAsync` à un sous-domaine du domaine du complément, cela ne fonctionnera pas. Le domaine complet et tous les sous-domaines doivent être exactement les mêmes.

Une fois que la première page (ou toute autre ressource) est chargée, un utilisateur peut utiliser des liens ou une autre interface utilisateur pour accéder à n’importe quel site web (ou n’importe quelle autre ressource) qui utilise le protocole HTTPS. Vous pouvez également concevoir la première page de façon à ce que l’utilisateur soit immédiatement redirigé vers un autre site.

Par défaut, la boîte de dialogue occupera 80 % de la hauteur et de la largeur de l’écran de l’appareil, mais vous pouvez définir des pourcentages différents en transmettant un objet de configuration à la méthode, comme indiqué dans l’exemple suivant.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Pour voir un exemple de complément qui effectue ce type d’action, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example). Pour plus d’exemples qui utilisent `displayDialogAsync`, consultez [Exemples](#samples).

Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)

> [!NOTE]
> Vous ne pouvez ouvrir qu’une seule boîte de dialogue à partir d’une fenêtre hôte. Toute tentative d’ouverture d’une autre boîte de dialogue génère une erreur. Par exemple, si un utilisateur ouvre une boîte de dialogue à partir d’un volet Office, il ne peut pas ouvrir une deuxième boîte de dialogue à partir d’une autre page dans le volet Office. Toutefois, quand une boîte de dialogue est ouverte à partir d’une [commande de complément](../design/add-in-commands.md), la commande ouvre un nouveau fichier HTML (mais invisible) chaque fois qu’elle est sélectionnée. Cela crée une nouvelle fenêtre hôte (invisible), afin que chaque fenêtre de ce type puisse lancer sa propre boîte de dialogue. Pour plus d’informations, reportez-vous à [Erreurs provenant de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Tirer parti d’une option de performances dans Office sur le web

La propriété `displayInIframe` est une propriété supplémentaire dans l’objet de configuration que vous pouvez transmettre à `displayDialogAsync`. Lorsque cette propriété est définie sur `true` et que le complément est en cours d’exécution dans un document ouvert dans Office sur le web, la boîte de dialogue s’ouvre sous la forme d’un iframe flottant et non d’une fenêtre indépendante ; elle s’ouvre ainsi plus rapidement. Voici un exemple.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

La valeur par défaut est `false`, ce qui revient à omettre entièrement la propriété. Si le complément n’est pas en cours d’exécution dans Office sur le Web, le `displayInIframe` est ignoré.

> [!NOTE]
> Vous **ne devez pas** utiliser `displayInIframe: true` si la boîte de dialogue redirige à tout moment vers une page qui ne peut pas être ouverte dans un iframe. Par exemple, les pages de connexion de nombreux services web populaires, tels que google et compte Microsoft, ne peuvent pas être ouvertes dans un iframe.

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Envoi d’informations à la page hôte à partir de la boîte de dialogue

> [!NOTE]
>
> - Pour plus de clarté, dans cette section, nous appelons le message cible la *page* hôte, mais à proprement parler, les messages sont envoyés au [runtime](../testing/runtimes.md) dans le volet Office (ou au runtime qui héberge un [fichier de fonction](/javascript/api/manifest/functionfile)). La distinction n’est significative que dans le cas de la messagerie inter-domaines. Pour plus d’informations, consultez [Messagerie inter-domaines au runtime hôte](#cross-domain-messaging-to-the-host-runtime).
> - La boîte de dialogue ne peut pas communiquer avec la page hôte dans le volet Office, sauf si la bibliothèque d’API JavaScript Office est chargée dans la page. (Comme toute page qui utilise la bibliothèque d’API JavaScript Office, le script de la page doit initialiser le complément. Pour plus d’informations, voir [Initialiser votre complément Office](initialize-add-in.md).)

Le code dans la boîte de dialogue utilise la fonction [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) pour envoyer un message de chaîne à la page hôte. La chaîne peut être un mot, une phrase, un objet blob XML, un JSON stringifié ou tout autre élément qui peut être sérialisé en chaîne ou converti en chaîne. Voici un exemple.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
>
> - La `messageParent` fonction est l’une *des deux* SEULES API JS Office qui peuvent être appelées dans la boîte de dialogue.
> - L’autre API JS qui peut être appelée dans la boîte de dialogue est `Office.context.requirements.isSetSupported`. Pour plus d’informations à ce sujet, voir [Spécifier les applications Office et les exigences d’API](specify-office-hosts-and-api-requirements.md). Toutefois, dans la boîte de dialogue, cette API n’est pas prise en charge dans les Outlook 2016 perpétuels sous licence en volume (autrement dit, la version MSI).

Dans l’exemple suivant, `googleProfile` est une version convertie en chaîne du profil Google de l’utilisateur.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

La page hôte doit être configurée de façon à recevoir le message. Pour ce faire, ajoutez un paramètre de rappel à l’appel d’origine de `displayDialogAsync`. Le rappel attribue un gestionnaire à l’événement `DialogMessageReceived`. Voici un exemple.

```js
let dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
>
> - Office transmet un objet [AsyncResult](/javascript/api/office/office.asyncresult) au rappel. Il représente le résultat de la tentative d’ouverture de la boîte de dialogue. Il ne représente pas le résultat de tous les événements dans la boîte de dialogue. Pour plus d’informations sur cette distinction, consultez la [Gestion des erreurs et des événements](dialog-handle-errors-events.md).
> - La propriété `value` de `asyncResult` est définie sur un objet [Dialog](/javascript/api/office/office.dialog), qui existe dans la page hôte, pas dans le contexte d’exécution de la boîte de dialogue.
> - The `processMessage` is the function that handles the event. You can give it any name you want.
> - La variable `dialog` est déclarée avec une portée plus large que le rappel, car elle est également référencée dans `processMessage`.

Voici un exemple simple de gestionnaire pour l’événement `DialogMessageReceived`.

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
>
> - Office transmet l’objet `arg` au gestionnaire. Sa `message` propriété est la chaîne envoyée par l’appel de dans `messageParent` la boîte de dialogue. Dans cet exemple, il s’agit d’une représentation stringifiée du profil d’un utilisateur à partir d’un service tel que le compte Microsoft ou Google. Il est donc désérialisé en objet avec `JSON.parse`.
> - L’implémentation `showUserName` n’est pas affichée. Elle peut afficher un message de bienvenue personnalisé dans le volet Office.

Lorsque l’intervention de l’utilisateur sur la boîte de dialogue est terminée, votre gestionnaire de messages doit fermer la boîte de dialogue, comme indiqué dans cet exemple.

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
>
> - L’objet `dialog` doit être le même que celui renvoyé par l’appel de `displayDialogAsync`.
> - L’appel de `dialog.close` indique à Office de fermer immédiatement la boîte de dialogue.

Pour voir un exemple de complément qui utilise ces techniques, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example.

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

Pour voir un exemple de complément qui effectue ce type d’action, consultez l’article relatif à l’exemple [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

### <a name="conditional-messaging"></a>Messagerie conditionnelle

Étant donné que vous pouvez envoyer plusieurs appels `messageParent` à partir de la boîte de dialogue, mais que vous n’avez qu’un seul gestionnaire dans la page hôte pour l’événement `DialogMessageReceived`, le gestionnaire doit utiliser la logique conditionnelle pour distinguer les différents messages. Par exemple, si la boîte de dialogue invite un utilisateur à se connecter à un fournisseur d’identité tel qu’un compte Microsoft ou Google, elle envoie le profil de l’utilisateur sous forme de message. Si l’authentification échoue, la boîte de dialogue envoie des informations d’erreur à la page hôte, comme dans l’exemple suivant.

```js
if (loginSuccess) {
    const userProfile = getProfile();
    const messageObject = {messageType: "signinSuccess", profile: userProfile};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    const errorDetails = getError();
    const messageObject = {messageType: "signinFailure", error: errorDetails};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
>
> - La variable `loginSuccess` serait initialisée en lisant la réponse HTTP à partir du fournisseur d’identité.
> - L’implémentation des `getProfile` fonctions et n’est `getError` pas affichée. Chacune obtient des données à partir d’un paramètre de requête ou du corps de la réponse HTTP.
> - Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.

The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

> [!NOTE]
> L’implémentation `showNotification` n’est pas affichée dans l’exemple de code fourni par cet article. Pour un exemple d'implémentation de cette fonction dans votre complément, voir [Exemple d'API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

### <a name="cross-domain-messaging-to-the-host-runtime"></a>Messagerie inter-domaines vers le runtime hôte

Une fois la boîte de dialogue ouverte, la boîte de dialogue ou le runtime parent peut s’éloigner du domaine du complément. Si l’une de ces choses se produit, un appel de `messageParent` échoue, sauf si votre code spécifie le domaine du runtime parent. Pour ce faire, ajoutez un paramètre [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) à l’appel de `messageParent`. Cet objet a une `targetOrigin` propriété qui spécifie le domaine auquel le message doit être envoyé. Si le paramètre n’est pas utilisé, Office suppose que la cible est le même domaine que celui hébergé par la boîte de dialogue.

> [!NOTE]
> L’utilisation `messageParent` de pour envoyer un message inter-domaines nécessite [l’ensemble de conditions requises Dialog Origin 1.1](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets). Le `DialogMessageOptions` paramètre est ignoré sur les versions antérieures d’Office qui ne prennent pas en charge l’ensemble de conditions requises, de sorte que le comportement de la méthode n’est pas affecté si vous la transmettez.

Voici un exemple d’utilisation `messageParent` de pour envoyer un message inter-domaines.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "https://resource.contoso.com" });
```

> [!NOTE]
> Le `DialogMessageOptions` paramètre a été publié vers le 19 juillet 2021. Pendant environ 30 jours après cette date, dans Office sur le Web, la première fois qui est appelée sans le `DialogMessageOptions` paramètre et que `messageParent` le parent est un domaine différent de la boîte de dialogue, l’utilisateur est invité à approuver l’envoi de données au domaine cible. Si l’utilisateur approuve, la réponse de l’utilisateur est mise en cache pendant 24 heures. L’utilisateur n’est pas invité à nouveau pendant cette période quand `messageParent` est appelé avec le même domaine cible.

Si le message n’inclut pas de données sensibles, vous pouvez définir sur `targetOrigin` «\* », ce qui permet de l’envoyer à n’importe quel domaine. Voici un exemple.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "*" });
```

> [!TIP]
> Le `DialogMessageOptions` paramètre a été ajouté à la `messageParent` méthode en tant que paramètre obligatoire à la mi-2021. Les compléments plus anciens qui envoient un message inter-domaines avec la méthode ne fonctionnent plus tant qu’ils n’ont pas été mis à jour pour utiliser le nouveau paramètre. Tant que le complément n’est pas mis à jour, *dans Office sur Windows uniquement*, les utilisateurs et les administrateurs système peuvent permettre à ces compléments de continuer à fonctionner en spécifiant les domaines approuvés avec un paramètre de Registre : **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Pour ce faire, créez un fichier avec une `.reg` extension, enregistrez-le sur l’ordinateur Windows, puis double-cliquez dessus pour l’exécuter. Voici un exemple du contenu d’un tel fichier.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="pass-information-to-the-dialog-box"></a>Transmission d’informations à la boîte de dialogue

Votre complément peut envoyer des messages de la [page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) à une boîte de dialogue à l’aide de [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)).

### <a name="use-messagechild-from-the-host-page"></a>Utiliser `messageChild()` à partir de la page hôte

Lorsque vous appelez l’API de boîte de dialogue Office pour ouvrir une boîte de dialogue, un objet [Dialog](/javascript/api/office/office.dialog) est retourné. Il doit être affecté à une variable dont l’étendue est supérieure à celle de la méthode [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) , car l’objet sera référencé par d’autres méthodes. Voici un exemple.

```javascript
let dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

Cet `Dialog` objet a une méthode [messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) qui envoie toute chaîne, y compris les données stringifiées, à la boîte de dialogue. Cela déclenche un `DialogParentMessageReceived` événement dans la boîte de dialogue. Votre code doit gérer cet événement, comme indiqué dans la section suivante.

Imaginez un scénario dans lequel l’interface utilisateur de la boîte de dialogue est liée à la feuille de calcul active et à la position de cette feuille de calcul par rapport aux autres feuilles de calcul. Dans l’exemple suivant, `sheetPropertiesChanged` envoie les propriétés de la feuille de calcul Excel à la boîte de dialogue. Dans ce cas, la feuille de calcul active est nommée « Ma feuille » et il s’agit de la deuxième feuille du classeur. Les données sont encapsulées dans un objet et stringifiées afin qu’elles puissent être passées à `messageChild`.

```javascript
function sheetPropertiesChanged() {
    const messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Gérer DialogParentMessageReceived dans la boîte de dialogue

Dans le code JavaScript de la boîte de dialogue, inscrivez un gestionnaire pour l’événement `DialogParentMessageReceived` avec la méthode [UI.addHandlerAsync](/javascript/api/office/office.ui#office-office-ui-addhandlerasync-member(1)) . Cette opération est généralement effectuée dans la [fonction Office.onReady ou Office.initialize](initialize-add-in.md), comme indiqué ci-dessous. (Un exemple plus robuste est inclus plus loin dans cet article.)

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Ensuite, définissez le `onMessageFromParent` gestionnaire. Le code suivant poursuit l’exemple de la section précédente. Notez qu’Office transmet un argument au gestionnaire et que la `message` propriété de l’objet argument contient la chaîne de la page hôte. Dans cet exemple, le message est reconverti en objet et jQuery est utilisé pour définir le titre supérieur de la boîte de dialogue pour qu’il corresponde au nouveau nom de feuille de calcul.

```javascript
function onMessageFromParent(arg) {
    const messageFromParent = JSON.parse(arg.message);
    $('h1').text(messageFromParent.name);
}
```

Il est recommandé de vérifier que votre gestionnaire est correctement inscrit. Pour ce faire, passez un rappel à la `addHandlerAsync` méthode . Cette opération s’exécute lorsque la tentative d’inscription du gestionnaire est terminée. Utilisez le gestionnaire pour journaliser ou afficher une erreur si le gestionnaire n’a pas été correctement inscrit. Voici un exemple. Notez que `reportError` est une fonction, non définie ici, qui journalise ou affiche l’erreur.

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a>Messagerie conditionnelle de la page parente vers la boîte de dialogue

Étant donné que vous pouvez effectuer plusieurs `messageChild` appels à partir de la page hôte, mais que vous n’avez qu’un seul gestionnaire dans la boîte de dialogue pour l’événement `DialogParentMessageReceived` , le gestionnaire doit utiliser la logique conditionnelle pour distinguer les différents messages. Vous pouvez le faire d’une manière qui est précisément parallèle à la façon dont vous structurez la messagerie conditionnelle lorsque la boîte de dialogue envoie un message à la page hôte, comme décrit dans [Messagerie conditionnelle](#conditional-messaging).

> [!NOTE]
> Dans certains cas, l’API `messageChild` , qui fait partie de [l’ensemble de conditions requises DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), peut ne pas être prise en charge. D’autres méthodes de messagerie parent-à-boîte de dialogue sont décrites dans [Autres façons de transmettre des messages à une boîte de dialogue à partir de sa page hôte](parent-to-dialog.md).

> [!IMPORTANT]
> [L’ensemble de conditions requises DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets) ne peut pas être spécifié dans la **\<Requirements\>** section d’un manifeste de complément. Vous devez vérifier la prise en charge de DialogApi 1.2 lors de l’exécution à l’aide de la `isSetSupported` méthode décrite dans Vérifications de l’exécution [pour la prise en charge de la méthode et de l’ensemble de conditions requises](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). La prise en charge des exigences de manifeste est en cours de développement.

### <a name="cross-domain-messaging-to-the-dialog-runtime"></a>Messagerie inter-domaines vers le runtime de dialogue

Une fois la boîte de dialogue ouverte, la boîte de dialogue ou le runtime parent peut s’éloigner du domaine du complément. Si l’une de ces choses se produit, les appels à `messageChild` échouent, sauf si votre code spécifie le domaine du runtime de la boîte de dialogue. Pour ce faire, ajoutez un paramètre [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) à l’appel de `messageChild`. Cet objet a une `targetOrigin` propriété qui spécifie le domaine auquel le message doit être envoyé. Si le paramètre n’est pas utilisé, Office suppose que la cible est le même domaine que celui que le runtime parent héberge actuellement.

> [!NOTE]
> L’utilisation `messageChild` de pour envoyer un message inter-domaines nécessite [l’ensemble de conditions requises Dialog Origin 1.1](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets). Le `DialogMessageOptions` paramètre est ignoré sur les versions antérieures d’Office qui ne prennent pas en charge l’ensemble de conditions requises, de sorte que le comportement de la méthode n’est pas affecté si vous la transmettez.

Voici un exemple d’utilisation `messageChild` de pour envoyer un message inter-domaines.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "https://resource.contoso.com" });
```

Si le message n’inclut pas de données sensibles, vous pouvez définir sur `targetOrigin` «\* », ce qui permet de *l’envoyer* à n’importe quel domaine. Voici un exemple.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "*" });
```

Étant donné que le runtime qui héberge la boîte de dialogue ne peut pas accéder à la **\<AppDomains\>** section du manifeste et déterminer ainsi si le domaine *d’où provient le message* est approuvé, vous devez utiliser le `DialogParentMessageReceived` gestionnaire pour le déterminer. L’objet passé au gestionnaire contient le domaine actuellement hébergé dans le parent en tant que `origin` propriété . Voici un exemple d’utilisation de la propriété .

```javascript
function onMessageFromParent(arg) {
    if (arg.origin === "https://addin.fabrikam.com") {
        // process message
    } else {
        dialog.close();
        showNotification("Messages from " + arg.origin + " are not accepted.");
    }
}
```

Par exemple, votre code peut utiliser la [fonction Office.onReady ou Office.initialize](initialize-add-in.md) pour stocker un tableau de domaines approuvés dans une variable globale. La `arg.origin` propriété peut ensuite être vérifiée par rapport à cette liste dans le gestionnaire.

> [!TIP]
> Le `DialogMessageOptions` paramètre a été ajouté à la `messageChild` méthode en tant que paramètre obligatoire à la mi-2021. Les compléments plus anciens qui envoient un message inter-domaines avec la méthode ne fonctionnent plus tant qu’ils n’ont pas été mis à jour pour utiliser le nouveau paramètre. Tant que le complément n’est pas mis à jour, *dans Office sur Windows uniquement*, les utilisateurs et les administrateurs système peuvent permettre à ces compléments de continuer à fonctionner en spécifiant les domaines approuvés avec un paramètre de Registre : **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Pour ce faire, créez un fichier avec une `.reg` extension, enregistrez-le sur l’ordinateur Windows, puis double-cliquez dessus pour l’exécuter. Voici un exemple du contenu d’un tel fichier.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="close-the-dialog-box"></a>Fermer la boîte de dialogue

You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example.

```js
function closeButtonClick() {
    const messageObject = {messageType: "dialogClosed"};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

Le gestionnaire de la page hôte pour `DialogMessageReceived` appelle `dialog.close`, comme dans cet exemple. (consultez les exemples précédents qui montrent comment l’objet `dialog` est initialisé).

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

Même lorsque vous ne disposez pas de votre propre interface utilisateur de fermeture de boîte de dialogue, un utilisateur final peut fermer la boîte de dialogue en choisissant le **X** dans le coin supérieur droit. Cette action déclenche l’événement `DialogEventReceived`. Si votre volet hôte a besoin de savoir quand cela se produit, il doit déclarer un gestionnaire pour cet événement. Pour plus d’informations, consultez la section [Erreurs et événements dans la boîte de dialogue](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box).

## <a name="advanced-topics-and-special-scenarios"></a>Rubriques plus complexes et scénarios spéciaux

### <a name="use-the-dialog-api-to-show-a-video"></a>Utilisation d'un API de boîte de dialogue pour afficher une vidéo

Voir [Utiliser la boîte de dialogue Office pour afficher une vidéo](dialog-video.md).

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a>Utilisation des API de boîte de dialogue dans un flux d’authentification

Voir [Authentifier avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).

### <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Utiliser l’API de boîte de dialogue Office avec des applications monopages et un routage côté client

Les authentifications par mot de passe (SPA) et le routage client doivent être gérés avec précaution lorsque vous utilisez l’API de boîte de dialogue Office. Consultez les [Pratiques recommandées pour l’utilisation de l’API de boîte de dialogue Office dans une SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).

### <a name="error-and-event-handling"></a>Gestion d'erreurs et d'événements

Voir [Gestion des erreurs et des événements dans la boîte de dialogue Office](dialog-handle-errors-events.md).

## <a name="next-steps"></a>Étapes suivantes

Découvrez les pièges et pratiques recommandées pour l’API de boîte de dialogue Office dans les [Meilleures pratiques et règles pour l’API de boîte de dialogue Office](dialog-best-practices.md).

## <a name="samples"></a>Échantillons

Tous les exemples suivants utilisent `displayDialogAsync`. Certains ont des serveurs basés sur NodeJS et d’autres ont des serveurs ASP.NET/IIS-based, mais la logique d’utilisation de la méthode est la même, quelle que soit la façon dont le complément est implémenté côté serveur.

**Bases:**

- [Exemple d’API de dialogue pour compléments Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [Contenu de formation / Création de compléments (plusieurs exemples)](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**Exemples plus complexes :**

- [Complément Office Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Complément Office Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [SSO NodeJS pour complément Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
- [Authentification unique ASPNET du complément Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Exemple de monétisation SAAS de complément Office](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [Complément Outlook Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Authentification unique de complément Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Visionneuse de jetons de complément Outlook](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Message actionnable du complément Outlook](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [Partage de complément Outlook sur OneDrive](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [Complément PowerPoint Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Scénario d’exécution partagé Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario)
- [Complément Excel ASPNET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Rédaction du complément Word JS](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Complément Word JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [Complément Word AngularJS Client OAuth](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Complément Office dans Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Complément Office OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Code des modèles de conception d’expérience utilisateur des compléments Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

** Voir aussi**

- [Runtimes dans les compléments Office](../testing/runtimes.md)