---
title: Utiliser l’API de boîte de dialogue Office dans vos compléments Office
description: Découvrir les notions de base relatives à la création d’une boîte de dialogue dans un complément Office
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: ed77173f57c8a16344d469585610917a08d3dcad
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324679"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Utiliser l’API de boîte de dialogue Office dans les compléments Office

Vous pouvez utiliser l’[API de dialogue Office](/javascript/api/office/office.ui) pour ouvrir des boîtes de dialogue dans votre complément Office. Cet article fournit des conseils concernant l’utilisation de l’API de dialogue dans votre complément Office.

> [!NOTE]
> Pour plus d’informations sur les compléments où l’API de dialogue est actuellement prise en charge, consultez la rubrique relative aux [ensembles de conditions requises de l’API de dialogue](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets). L’API de dialogue est actuellement prise en charge pour Word, Excel, PowerPoint et Outlook.

Un scénario principal pour l’API de dialogue consiste à activer l’authentification à l'aide d'une ressource telle que Google, Facebook, ou Microsoft Graph. Pour plus d’informations, voir [S’authentifier auprès de l'API de boîte de dialogue Office](auth-with-office-dialog-api.md) *une fois* que vous êtes familiarisé(e) avec cet article.

Envisagez d’ouvrir une boîte de dialogue à partir d’un volet Office, d’un complément de contenu ou d’un [complément de commande](../design/add-in-commands.md) pour effectuer les opérations suivantes :

- afficher les pages de connexion qui ne peuvent pas être ouvertes directement dans un volet Office ;
- fournir davantage d’espace à l’écran, ou même un plein écran, pour certaines tâches exécutées dans votre complément ;
- héberger une vidéo qui serait trop petite si elle était limitée à un volet Office.

> [!NOTE]
> Comme des éléments d’interface utilisateur qui se chevauchent peuvent gêner des utilisateurs, évitez d’ouvrir une boîte de dialogue à partir d’un volet Office à moins que votre scénario l’exige. Lorsque vous envisagez d’utiliser la surface d’exposition d’un volet Office, tenez compte du fait que les volets Office peuvent être affichés sous forme d’onglets. Pour voir un exemple, consultez la rubrique relative à l’exemple [Complément Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

L’image suivante montre un exemple de boîte de dialogue.

![Commandes de complément](../images/auth-o-dialog-open.png)

Notez que la boîte de dialogue s’ouvre toujours au centre de l’écran. L’utilisateur peut la déplacer et la redimensionner. La fenêtre est *non modale* : un utilisateur peut continuer à interagir à la fois avec le document dans l’application Office hôte et avec la page dans le volet Office, le cas échéant.

## <a name="open-a-dialog-box-from-a-host-page"></a>Ouvrir une boîte de dialogue à partir d’une page hôte

Les API JavaScript Office incluent un objet [Dialog](/javascript/api/office/office.dialog) et deux fonctions dans l’[espace de noms Office.context.ui](/javascript/api/office/office.ui).

Pour ouvrir une boîte de dialogue, généralement une page dans un volet des tâches, votre code appelle la méthode [displayDialogAsync](/javascript/api/office/office.ui) et lui transmet l’URL de la ressource que vous voulez ouvrir. La page sur laquelle cette méthode est appelée est connue sous le nom de « page hôte ». Par exemple, si vous appelez cette méthode dans le script sur index.html d'un volet de tâches, la page index.html correspond à la page hôte de la boîte de dialogue ouverte par la méthode.

La ressource ouverte dans la boîte de dialogue correspond généralement à une page, mais ce peut être une méthode du contrôleur dans une application MVC, un itinéraire, une méthode de service web ou toute autre ressource. Dans cet article, les termes « page » ou « site web » font référence à la ressource dans la boîte de dialogue. Le code suivant est un exemple simple :

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - L’URL utilise le protocole HTTP**S**. Ceci est obligatoire pour toutes les pages chargées dans une boîte de dialogue, pas seulement la première page chargée.
> - Le domaine de la boîte de dialogue est le même que celui de la page hôte, qui peut être la page d’un volet Office ou le [fichier de fonctions](/office/dev/add-ins/reference/manifest/functionfile) d’une commande de complément. Obligatoire : la page, la méthode du contrôleur ou toute autre ressource qui est transmise à la méthode `displayDialogAsync` doit se trouver dans le même domaine que la page hôte.

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
> Vous ne devez **pas** utiliser `displayInIframe: true` si la boîte de dialogue redirige à un moment donné l’utilisateur vers une page qui ne peut pas être ouverte dans un IFrame. Par exemple, les pages de connexion de nombreux services web connus, comme un compte Microsoft et Google, ne peuvent pas être ouvertes dans un IFrame.

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Envoi d’informations à la page hôte à partir de la boîte de dialogue

La boîte de dialogue ne peut pas communiquer avec la page hôte dans le volet Office, sauf si :

- la page active dans la boîte de dialogue se trouve dans le même domaine que la page hôte ;
- La bibliothèque de l’API JavaScript pour Office est chargée dans la page. (Comme n’importe quelle page qui utilise la bibliothèque d’API JavaScript d’Office, le script de la page doit `Office.initialize` assigner une méthode à la propriété, bien qu’il puisse s’agir d’une méthode vide. Pour plus d’informations, consultez [la rubrique initialiser votre complément Office](initialize-add-in.md).

Le code de la boîte de dialogue utilise la fonction [messageParent](/javascript/api/office/office.ui#messageparent-message-) pour envoyer une valeur booléenne ou un message de type chaîne à la page hôte. La chaîne peut être un mot, une phrase, un blob XML, un JSON converti en chaîne ou un autre élément pouvant être sérialisé en chaîne. Voici un exemple :

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - La fonction `messageParent` est l’une des deux *seules* API Office pouvant être appelées dans la boîte de dialogue. L’autre est `Office.context.requirements.isSetSupported`. Pour plus d’informations, consultez la rubrique relative à la [spécification d’hôtes Office et de conditions requises d’API](specify-office-hosts-and-api-requirements.md).
> - La fonction `messageParent` peut uniquement être appelée sur une page ayant le même domaine (y compris les mêmes protocole et port) que la page hôte.

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
> - Office transmet l’objet `arg` au gestionnaire. Sa propriété `message` est la valeur booléenne ou la chaîne envoyée par l’appel de `messageParent` dans la boîte de dialogue. Dans cet exemple, il s’agit d’une représentation convertie en chaîne du profil de l’utilisateur à partir d’un service tel qu’un compte Microsoft ou Google, qui est donc désérialisé en objet avec `JSON.parse`.
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

Étant donné que vous pouvez envoyer plusieurs appels `messageParent` à partir de la boîte de dialogue, mais que vous n’avez qu’un seul gestionnaire dans la page hôte pour l’événement `DialogMessageReceived`, le gestionnaire doit utiliser la logique conditionnelle pour distinguer les différents messages. Par exemple, si la boîte de dialogue invite un utilisateur à se connecter à un fournisseur d’identité tel qu’un compte Microsoft ou Google, elle envoie le profil de l’utilisateur sous la forme d’un message. Si l’authentification échoue, la boîte de dialogue envoie des informations sur l’erreur à la page hôte, comme dans l’exemple suivant :

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

Parfois, la page hôte doit transmettre des informations à la boîte de dialogue. Pour ce faire, il existe deux moyens :

- ajouter des paramètres de requête à l’URL qui est transmise à `displayDialogAsync` ;
- stocker les informations à un emplacement auquel à la fois la fenêtre hôte et la boîte de dialogue ont accès. Les deux fenêtres ne partagent pas un stockage de session commun, mais *si elles ont le même domaine* (y compris le même numéro de port, le cas échéant), elles utilisent un [Stockage local](https://www.w3schools.com/html/html5_webstorage.asp) commun.\*

> [!NOTE]
> \* Un bogue peut affecter votre stratégie de gestion des jetons. Si le complément s’exécute dans **Office sur le web** dans le navigateur Safari ou Edge, la boîte de dialogue et le volet des tâches Office ne partagent pas le même stockage local. Il ne peut donc pas être utilisé pour communiquer entre eux.

### <a name="use-local-storage"></a>Utilisation du stockage local

Pour utiliser le stockage local, votre code appelle la méthode `setItem` de l’objet `window.localStorage` dans la page hôte avant l’appel de `displayDialogAsync`, comme dans l’exemple suivant :

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Le code dans la boîte de dialogue qui lit l’élément lorsqu’il est nécessaire, comme dans l’exemple suivant :

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

### <a name="use-query-parameters"></a>Utiliser les paramètres de requête

L’exemple suivant montre comment transmettre des données à l’aide d’un paramètre de requête :

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Pour obtenir un exemple qui utilise cette technique, consultez l’article relatif à l’exemple [Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Le code dans votre boîte de dialogue peut analyser l’URL et lire la valeur du paramètre.

> [!NOTE]
> Office ajoute automatiquement un paramètre de requête appelé `_host_info` à l’URL qui est transmise à `displayDialogAsync`. (Il est ajouté après vos paramètres de requête personnalisés, le cas échéant. Il n’est pas ajouté à toutes les autres URL auxquelles la boîte de dialogue accède.) Microsoft peut modifier le contenu de cette valeur, ou le supprimer entièrement, à l’avenir, donc votre code ne doit pas le lire. La même valeur est ajoutée au stockage de session de la boîte de dialogue. Là encore, *votre code ne doit ni lire, ni écrire cette valeur*.

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