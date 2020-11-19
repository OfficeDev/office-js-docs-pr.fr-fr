---
title: Utiliser l’API de boîte de dialogue Office dans vos compléments Office
description: Découvrez les concepts de base de la création d’une boîte de dialogue dans un complément Office.
ms.date: 10/21/2020
localization_priority: Normal
ms.openlocfilehash: 56c12aa5b15f8f79273923402c5a5bfa92a2dde8
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131758"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Utiliser l’API de boîte de dialogue Office dans les compléments Office

Vous pouvez utiliser l’[API de dialogue Office](/javascript/api/office/office.ui) pour ouvrir des boîtes de dialogue dans votre complément Office. Cet article fournit des conseils concernant l’utilisation de l’API de dialogue dans votre complément Office.

> [!NOTE]
> Pour plus d’informations sur les compléments où l’API de dialogue est actuellement prise en charge, consultez la rubrique relative aux [ensembles de conditions requises de l’API de dialogue](../reference/requirement-sets/dialog-api-requirement-sets.md). L’API de dialogue est actuellement prise en charge pour Excel, PowerPoint et Word. La prise en charge d’Outlook est incluse dans différents ensembles de conditions requises pour les boîtes aux lettres &mdash; consultez la référence de l’API pour plus de détails.

Un scénario principal pour l’API de dialogue consiste à activer l’authentification à l'aide d'une ressource telle que Google, Facebook, ou Microsoft Graph. Pour plus d’informations, voir [S’authentifier auprès de l'API de boîte de dialogue Office](auth-with-office-dialog-api.md) *une fois* que vous êtes familiarisé(e) avec cet article.

Envisagez d’ouvrir une boîte de dialogue à partir d’un volet Office, d’un complément de contenu ou d’un [complément de commande](../design/add-in-commands.md) pour effectuer les opérations suivantes :

- afficher les pages de connexion qui ne peuvent pas être ouvertes directement dans un volet Office ;
- fournir davantage d’espace à l’écran, ou même un plein écran, pour certaines tâches exécutées dans votre complément ;
- héberger une vidéo qui serait trop petite si elle était limitée à un volet Office.

> [!NOTE]
> Comme des éléments d’interface utilisateur qui se chevauchent peuvent gêner des utilisateurs, évitez d’ouvrir une boîte de dialogue à partir d’un volet Office à moins que votre scénario l’exige. Lorsque vous envisagez d’utiliser la surface d’exposition d’un volet Office, tenez compte du fait que les volets Office peuvent être affichés sous forme d’onglets. Pour voir un exemple, consultez la rubrique relative à l’exemple [Complément Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

L’image suivante montre un exemple de boîte de dialogue.

![Capture d’écran montrant la boîte de dialogue avec 3 options de connexion affichées devant Word](../images/auth-o-dialog-open.png)

Notez que la boîte de dialogue s’ouvre toujours au centre de l’écran. L’utilisateur peut la déplacer et la redimensionner. La fenêtre n’est pas *modale*: un utilisateur peut continuer à interagir avec le document dans l’application Office et avec la page dans le volet Office, le cas échéant.

## <a name="open-a-dialog-box-from-a-host-page"></a>Ouvrir une boîte de dialogue à partir d’une page hôte

Les API JavaScript Office incluent un objet [Dialog](/javascript/api/office/office.dialog) et deux fonctions dans l’[espace de noms Office.context.ui](/javascript/api/office/office.ui).

Pour ouvrir une boîte de dialogue, généralement une page dans un volet des tâches, votre code appelle la méthode [displayDialogAsync](/javascript/api/office/office.ui) et lui transmet l’URL de la ressource que vous voulez ouvrir. La page sur laquelle cette méthode est appelée est connue sous le nom de « page hôte ». Par exemple, si vous appelez cette méthode dans le script sur index.html d'un volet de tâches, la page index.html correspond à la page hôte de la boîte de dialogue ouverte par la méthode.

La ressource ouverte dans la boîte de dialogue correspond généralement à une page, mais ce peut être une méthode du contrôleur dans une application MVC, un itinéraire, une méthode de service web ou toute autre ressource. Dans cet article, les termes « page » ou « site web » font référence à la ressource dans la boîte de dialogue. Le code suivant est un exemple simple :

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - L’URL utilise le protocole HTTP **S**. Ceci est obligatoire pour toutes les pages chargées dans une boîte de dialogue, pas seulement la première page chargée.
> - Le domaine de la boîte de dialogue est le même que celui de la page hôte, qui peut être la page d’un volet Office ou le [fichier de fonctions](../reference/manifest/functionfile.md) d’une commande de complément. Obligatoire : la page, la méthode du contrôleur ou toute autre ressource qui est transmise à la méthode `displayDialogAsync` doit se trouver dans le même domaine que la page hôte.

> [!IMPORTANT]
> La page hôte et les ressources s'ouvrant dans la boîte de dialogue doivent avoir le même domaine complet. Si vous tentez de transmettre `displayDialogAsync` à un sous-domaine du domaine du complément, cela ne fonctionnera pas. Le domaine complet et tous les sous-domaines doivent être exactement les mêmes.

Une fois que la première page (ou toute autre ressource) est chargée, un utilisateur peut utiliser des liens ou une autre interface utilisateur pour accéder à n’importe quel site web (ou n’importe quelle autre ressource) qui utilise le protocole HTTPS. Vous pouvez également concevoir la première page de façon à ce que l’utilisateur soit immédiatement redirigé vers un autre site.

Par défaut, la boîte de dialogue occupera 80 % de la hauteur et de la largeur de l’écran de l’appareil, mais vous pouvez définir des pourcentages différents en transmettant un objet de configuration à la méthode, comme indiqué dans l’exemple suivant :

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Pour voir un exemple de complément qui effectue ce type d’action, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Définissez les deux valeurs sur 100 % pour bénéficier d’une réelle d’expérience de plein écran. (Le maximum réel est de 99,5 %, et la fenêtre peut toujours être déplacée et redimensionnée.)

> [!NOTE]
> Vous ne pouvez ouvrir qu’une seule boîte de dialogue à partir d’une fenêtre hôte. Toute tentative d’ouverture d’une autre boîte de dialogue génère une erreur. Par exemple, si un utilisateur ouvre une boîte de dialogue à partir d’un volet Office, il ne peut pas ouvrir une seconde boîte de dialogue à partir d’une autre page dans le volet Office. Toutefois, quand une boîte de dialogue est ouverte à partir d’une [commande de complément](../design/add-in-commands.md), la commande ouvre un nouveau fichier HTML (mais invisible) chaque fois qu’elle est sélectionnée. Cela crée une nouvelle fenêtre hôte (invisible), afin que chaque fenêtre de ce type puisse lancer sa propre boîte de dialogue. Pour plus d’informations, reportez-vous à [Erreurs provenant de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Tirer parti d’une option de performances dans Office sur le web

La propriété `displayInIframe` est une propriété supplémentaire dans l’objet de configuration que vous pouvez transmettre à `displayDialogAsync`. Lorsque cette propriété est définie sur `true` et que le complément est en cours d’exécution dans un document ouvert dans Office sur le web, la boîte de dialogue s’ouvre sous la forme d’un iframe flottant et non d’une fenêtre indépendante ; elle s’ouvre ainsi plus rapidement. Voici un exemple :

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

La valeur par défaut est `false`, ce qui revient à omettre entièrement la propriété. Si le complément n’est pas exécuté dans Office sur le Web, le `displayInIframe` est ignoré.

> [!NOTE]
> Vous ne devez **pas** utiliser `displayInIframe: true` si la boîte de dialogue redirige à un moment donné l’utilisateur vers une page qui ne peut pas être ouverte dans un IFrame. Par exemple, les pages de connexion de nombreux services Web populaires, tels que Google et Microsoft Account, ne peuvent pas être ouverts dans un IFRAME.

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Envoi d’informations à la page hôte à partir de la boîte de dialogue

La boîte de dialogue ne peut pas communiquer avec la page hôte dans le volet Office, sauf si :

- la page active dans la boîte de dialogue se trouve dans le même domaine que la page hôte ;
- La bibliothèque de l’API JavaScript pour Office est chargée dans la page. (Comme n’importe quelle page qui utilise la bibliothèque d’API JavaScript d’Office, le script de la page doit assigner une méthode à la `Office.initialize` propriété, bien qu’il puisse s’agir d’une méthode vide. Pour plus d’informations, consultez [la rubrique initialiser votre complément Office](initialize-add-in.md).

Le code de la boîte de dialogue utilise la fonction [messageParent](/javascript/api/office/office.ui#messageparent-message-) pour envoyer une valeur booléenne ou un message de type chaîne à la page hôte. La chaîne peut être un mot, une phrase, un blob XML, un JSON converti en chaîne ou un autre élément pouvant être sérialisé en chaîne. Voici un exemple :

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - La fonction `messageParent` peut uniquement être appelée sur une page ayant le même domaine (y compris les mêmes protocole et port) que la page hôte.
> - La `messageParent` fonction est l’une des deux *seules* deux API Office js qui peuvent être appelées dans la boîte de dialogue.
> - L’autre API JS qui peut être appelée dans la boîte de dialogue est `Office.context.requirements.isSetSupported` . Pour plus d’informations à ce sujet, voir [spécifier les applications Office et les conditions requises](specify-office-hosts-and-api-requirements.md)de l’API. Toutefois, dans la boîte de dialogue, cette API n’est pas prise en charge dans l’achat d’Outlook 2016 1 à l’heure (autrement dit, la version MSI).

Dans l’exemple suivant, `googleProfile` est une version convertie en chaîne du profil Google de l’utilisateur.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

La page hôte doit être configurée de façon à recevoir le message. Pour ce faire, ajoutez un paramètre de rappel à l’appel d’origine de `displayDialogAsync`. Le rappel attribue un gestionnaire à l’événement `DialogMessageReceived`. Voici un exemple :

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
> - Office transmet un objet [AsyncResult](/javascript/api/office/office.asyncresult) au rappel. Il représente le résultat de la tentative d’ouverture de la boîte de dialogue. Il ne représente pas le résultat de tous les événements dans la boîte de dialogue. Pour plus d’informations sur cette distinction, consultez la [Gestion des erreurs et des événements](dialog-handle-errors-events.md).
> - La propriété `value` de `asyncResult` est définie sur un objet [Dialog](/javascript/api/office/office.dialog), qui existe dans la page hôte, pas dans le contexte d’exécution de la boîte de dialogue.
> - `processMessage` est la fonction qui gère l’événement. Vous pouvez lui donner le nom que vous souhaitez.
> - La variable `dialog` est déclarée avec une portée plus large que le rappel, car elle est également référencée dans `processMessage`.

Voici un exemple simple de gestionnaire pour l’événement `DialogMessageReceived` :

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - Office transmet l’objet `arg` au gestionnaire. Sa propriété `message` est la valeur booléenne ou la chaîne envoyée par l’appel de `messageParent` dans la boîte de dialogue. Dans cet exemple, il s’agit d’une représentation JSON du profil d’un utilisateur à partir d’un service tel que le compte Microsoft ou Google, de sorte qu’il soit désérialisé en un objet avec `JSON.parse` .
> - L’implémentation `showUserName` n’est pas visible. Elle peut afficher un message de bienvenue personnalisé dans le volet Office.

Lorsque l’intervention de l’utilisateur sur la boîte de dialogue est terminée, votre gestionnaire de messages doit fermer la boîte de dialogue, comme indiqué dans cet exemple.

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - L’objet `dialog` doit être le même que celui renvoyé par l’appel de `displayDialogAsync`.
> - L’appel de `dialog.close` indique à Office de fermer immédiatement la boîte de dialogue.

Pour voir un exemple de complément qui utilise ces techniques, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Si le complément a besoin d’ouvrir une autre page du volet Office après avoir reçu le message, vous pouvez utiliser la méthode `window.location.replace` (ou `window.location.href`) en tant que dernière ligne du gestionnaire. Voici un exemple :

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

Étant donné que vous pouvez envoyer plusieurs appels `messageParent` à partir de la boîte de dialogue, mais que vous n’avez qu’un seul gestionnaire dans la page hôte pour l’événement `DialogMessageReceived`, le gestionnaire doit utiliser la logique conditionnelle pour distinguer les différents messages. Par exemple, si la boîte de dialogue invite l’utilisateur à se connecter à un fournisseur d’identité tel que le compte Microsoft ou Google, il envoie le profil de l’utilisateur sous la forme d’un message. Si l’authentification échoue, la boîte de dialogue envoie des informations sur l’erreur à la page hôte, comme dans l’exemple suivant :

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
> - La variable `loginSuccess` serait initialisée en lisant la réponse HTTP à partir du fournisseur d’identité.
> - L’implémentation des fonctions `getProfile` et `getError` n’est pas affichée. Chacune obtient des données à partir d’un paramètre de requête ou du corps de la réponse HTTP.
> - Des objets anonymes de différents types sont envoyés selon que la connexion a réussi ou non. Tous deux ont une propriété `messageType`, mais un a une propriété `profile` et l’autre une propriété `error`.

Le code du gestionnaire dans la page hôte utilise la valeur de la propriété `messageType` pour créer une branche comme le montre l’exemple suivant. Notez que la fonction `showUserName` est identique à celle de l’exemple précédent et que la fonction `showNotification` affiche l’erreur dans l’interface utilisateur de la page hôte.

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
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
> L'implémentation `showNotification` n'est pas montrée dans l'exemple de code fourni par cet article. Pour un exemple d'implémentation de cette fonction dans votre complément, voir [Exemple d'API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

## <a name="pass-information-to-the-dialog-box"></a>Transmission d’informations à la boîte de dialogue

Votre complément peut envoyer des messages à partir de la [page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) vers une boîte de dialogue à l’aide de la [boîte de dialogue. messageChild](/javascript/api/office/office.dialog#messagechild-message-).

### <a name="use-messagechild-from-the-host-page"></a>Utiliser `messageChild()` à partir de la page hôte

Lorsque vous appelez l’API de boîte de dialogue Office pour ouvrir une boîte de dialogue, un objet [Dialog](/javascript/api/office/office.dialog) est renvoyé. Elle doit être assignée à une variable dont l’étendue est supérieure à celle de la méthode [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) , car l’objet sera référencé par d’autres méthodes. Voici un exemple :

```javascript
var dialog;
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

Cet `Dialog` objet est doté d’une méthode [messageChild](/javascript/api/office/office.dialog#messagechild-message-) qui envoie n’importe quelle chaîne, y compris les données JSON, à la boîte de dialogue. Cela déclenche un `DialogParentMessageReceived` événement dans la boîte de dialogue. Votre code doit gérer cet événement, comme indiqué dans la section suivante.

Imaginez un scénario dans lequel l’interface utilisateur de la boîte de dialogue est liée à la feuille de calcul active et la position de cette feuille de calcul par rapport aux autres feuilles de calcul. Dans l’exemple suivant, `sheetPropertiesChanged` envoie les propriétés de feuille de calcul Excel dans la boîte de dialogue. Dans ce cas, la feuille de calcul active est nommée « ma feuille » et est la seconde feuille du classeur. Les données sont encapsulées dans un objet et JSON de sorte qu’il soit possible de les transmettre à `messageChild` .

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Gérer DialogParentMessageReceived dans la boîte de dialogue

Dans le JavaScript de la boîte de dialogue, inscrivez un gestionnaire pour l' `DialogParentMessageReceived` événement à l’aide de la méthode [UI. addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) . Cette opération s’effectue généralement dans les [méthodes Office. onReady ou Office.initialize](initialize-add-in.md), comme indiqué dans l’exemple suivant. (Un exemple plus fiable est ci-dessous.)

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Ensuite, définissez le `onMessageFromParent` Gestionnaire. Le code suivant poursuit l’exemple de la section précédente. Notez qu’Office transmet un argument au gestionnaire et que la `message` propriété de l’objet argument contient la chaîne de la page hôte. Dans cet exemple, le message est reconverti en objet et jQuery est utilisé pour définir le titre supérieur de la boîte de dialogue de sorte qu’il corresponde au nouveau nom de la feuille de calcul.

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

Il est recommandé de vérifier que votre gestionnaire est correctement enregistré. Vous pouvez effectuer cette opération en transférant un rappel à la `addHandlerAsync` méthode. Cette opération s’exécute lorsque la tentative d’enregistrement du gestionnaire est terminée. Utilisez le gestionnaire pour consigner ou afficher une erreur si le gestionnaire n’a pas été enregistré correctement. Voici un exemple. Notez qu' `reportError` il s’agit d’une fonction, non définie ici, qui enregistre ou affiche l’erreur.

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

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a>Messagerie conditionnelle de la page parent à la boîte de dialogue

Étant donné que vous pouvez effectuer plusieurs `messageChild` appels à partir de la page hôte, mais que vous n’avez qu’un seul gestionnaire dans la boîte de dialogue de l' `DialogParentMessageReceived` événement, le gestionnaire doit utiliser une logique conditionnelle pour distinguer les différents messages. Vous pouvez effectuer cette opération d’une manière parfaitement parallèle à la façon dont vous structurez la messagerie conditionnelle lorsque la boîte de dialogue envoie un message à la page hôte, comme décrit dans la section [messagerie conditionnelle](#conditional-messaging).

> [!NOTE]
> Dans certains cas, l' `messageChild` API, qui fait partie de l' [ensemble de conditions requises DialogApi 1,2](../reference/requirement-sets/dialog-api-requirement-sets.md), n’est peut-être pas prise en charge. D’autres méthodes pour la messagerie de type parent-à-boîte de dialogue sont décrites dans [les autres façons de transmettre des messages à une boîte de dialogue à partir de sa page hôte](parent-to-dialog.md).

> [!IMPORTANT]
> L' [ensemble de conditions requises DialogApi 1,2](../reference/requirement-sets/dialog-api-requirement-sets.md) ne peut pas être spécifié dans la `<Requirements>` section d’un manifeste de complément. Vous devrez vérifier la prise en charge de DialogApi 1,2 lors de l’exécution à l’aide de la méthode [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) . La prise en charge des exigences de manifeste est en cours de développement.

## <a name="closing-the-dialog-box"></a>Fermeture de la boîte de dialogue

Vous pouvez implémenter un bouton de fermeture dans la boîte de dialogue. Pour ce faire, le gestionnaire d’événements Click du bouton doit utiliser `messageParent` pour indiquer à la page hôte que vous avez cliqué sur le bouton. Voici un exemple :

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

Le gestionnaire de la page hôte pour `DialogMessageReceived` appelle `dialog.close`, comme dans cet exemple. (consultez les exemples précédents qui montrent comment l’objet `dialog` est initialisé).

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
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

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Utilisation de l’API de boîte de dialogue Office avec des applications à page unique et routage côté client

Les authentifications par mot de passe (SPA) et le routage client doivent être gérés avec précaution lorsque vous utilisez l’API de boîte de dialogue Office. Consultez les [Pratiques recommandées pour l’utilisation de l’API de boîte de dialogue Office dans une SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).

### <a name="error-and-event-handling"></a>Gestion d'erreurs et d'événements

Voir [Gestion des erreurs et des événements dans la boîte de dialogue Office](dialog-handle-errors-events.md).

## <a name="next-steps"></a>Étapes suivantes

Découvrez les pièges et pratiques recommandées pour l’API de boîte de dialogue Office dans les [Meilleures pratiques et règles pour l’API de boîte de dialogue Office](dialog-best-practices.md).
