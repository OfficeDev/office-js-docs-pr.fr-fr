---
title: Utiliser l?API de dialogue dans vos compl?ments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b026c3c5871372c52d0b44e36c01fc44a3d2bf04
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a>Utiliser l?API de dialogue dans vos compl?ments Office

Vous pouvez utiliser l?[API de dialogue](https://dev.office.com/reference/add-ins/shared/officeui) pour ouvrir des bo?tes de dialogue dans votre compl?ment Office. Cet article fournit des conseils concernant l?utilisation de l?API de dialogue dans votre compl?ment Office.

> [!NOTE]
> Pour plus d?informations sur les compl?ments o? l?API de dialogue est actuellement prise en charge, consultez la rubrique relative aux [ensembles de conditions requises de l?API de dialogue](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets). L?API de dialogue est actuellement prise en charge pour Word, Excel, PowerPoint et Outlook.

> Un sc?nario principal pour l?API de dialogue consiste ? activer l?authentification pour une ressource telle que Google ou Facebook. Si votre compl?ment n?cessite les donn?es relatives ? l?utilisateur d?Office ou leurs ressources accessibles via Microsoft Graph, par exemple Office 365 ou OneDrive, nous vous recommandons d?utiliser l?API d?authentification unique chaque fois que possible. Si vous utilisez les API pour l?authentification unique, vous n?aurez pas besoin de l?API de dialogue. Pour plus d?informations, consultez la rubrique [Activer l?authentification unique pour des compl?ments Office](sso-in-office-add-ins.md).

Envisagez d?ouvrir une bo?te de dialogue ? partir d?un volet Office, d?un compl?ment de contenu ou d?un [compl?ment de commande](../design/add-in-commands.md) pour effectuer les op?rations suivantes :

- afficher les pages de connexion qui ne peuvent pas ?tre ouvertes directement dans un volet Office ;
- fournir davantage d?espace ? l??cran, ou m?me un plein ?cran, pour certaines t?ches ex?cut?es dans votre compl?ment ;
- h?berger une vid?o qui serait trop petite si elle ?tait limit?e ? un volet Office.

> [!NOTE]
> Comme des ?l?ments d?IU qui se chevauchent peuvent g?ner des utilisateurs, ?vitez d?ouvrir une bo?te de dialogue ? partir d?un volet Office ? moins que votre sc?nario l?exige. Lorsque vous envisagez d?utiliser la surface d?exposition d?un volet Office, tenez compte du fait que les volets Office peuvent ?tre affich?s sous forme d?onglets. Pour voir un exemple, consultez la rubrique relative ? l?exemple de [compl?ment Excel JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

L?image suivante montre un exemple de bo?te de dialogue.

![Commandes de compl?ment](../images/auth-o-dialog-open.png)

Notez que la bo?te de dialogue s?ouvre toujours au centre de l??cran. L?utilisateur peut la d?placer et la redimensionner. La fen?tre est *non modale* : un utilisateur peut continuer ? interagir ? la fois avec le document dans l?application Office h?te et avec la page h?te dans le volet Office, le cas ?ch?ant.

## <a name="dialog-api-scenarios"></a>Sc?narios de l?API de dialogue

Les API JavaScript Office prennent en charge les sc?narios suivants avec un objet [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) et deux fonctions dans l?[espace de noms Office.context.ui](https://dev.office.com/reference/add-ins/shared/officeui).

### <a name="open-a-dialog-box"></a>Ouvrir une bo?te de dialogue.

Pour ouvrir une bo?te de dialogue, votre code dans le volet Office appelle la m?thode [displayDialogAsync](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) et lui transmet l?URL de la ressource que vous voulez ouvrir. Il s?agit g?n?ralement d?une page, mais ce peut ?tre une m?thode du contr?leur dans une application MVC, un itin?raire, une m?thode de service web ou toute autre ressource. Dans cet article, les termes ? page ? ou ? site web ? font r?f?rence ? la ressource dans la bo?te de dialogue. Le code suivant est un exemple simple.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - L?URL utilise le protocole HTTP**S**. Ceci est obligatoire pour toutes les pages charg?es dans une bo?te de dialogue, pas seulement la premi?re page charg?e.
> - Le domaine est le m?me que celui de la page h?te, qui peut ?tre la page d?un volet Office ou le [fichier de fonctions](https://dev.office.com/reference/add-ins/manifest/functionfile) d?une commande de compl?ment. Obligatoire : la page, la m?thode du contr?leur ou toute autre ressource qui est transmise ? la m?thode `displayDialogAsync` doit se trouver dans le m?me domaine que la page h?te.

Une fois que la premi?re page (ou toute autre ressource) est charg?e, un utilisateur peut acc?der ? n?importe quel site web (ou n?importe quelle autre ressource) qui utilise le protocole HTTPS. Vous pouvez ?galement concevoir la premi?re page de fa?on ? ce que l?utilisateur soit imm?diatement redirig? vers un autre site.

Par d?faut, la bo?te de dialogue occupera 80 % de la hauteur et de la largeur de l??cran de l?appareil, mais vous pouvez d?finir des pourcentages diff?rents en transmettant un objet de configuration ? la m?thode, comme indiqu? dans l?exemple suivant :

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Pour voir un exemple de compl?ment qui effectue ce type d?action, consultez la rubrique relative ? l?[exemple d?API de dialogue de compl?ment Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

D?finissez les deux valeurs sur 100 % pour b?n?ficier d?une r?elle d?exp?rience de plein ?cran. (Le maximum r?el est de 99,5 %, et la fen?tre peut toujours ?tre d?plac?e et redimensionn?e.)

> [!NOTE]
> Vous ne pouvez ouvrir qu?une seule bo?te de dialogue ? partir d?une fen?tre h?te. Toute tentative d?ouverture d?une autre bo?te de dialogue g?n?re une erreur. Par exemple, si un utilisateur ouvre une bo?te de dialogue ? partir d?un volet Office, il ne peut pas ouvrir une seconde bo?te de dialogue ? partir d?une autre page dans le volet Office. Toutefois, quand une bo?te de dialogue est ouverte ? partir d?une [commande de compl?ment](../design/add-in-commands.md), la commande ouvre un nouveau fichier HTML (mais invisible) chaque fois qu?elle est s?lectionn?e. Cela cr?e une nouvelle fen?tre h?te (invisible), afin que chaque fen?tre de ce type puisse lancer sa propre bo?te de dialogue. Pour plus d?informations, reportez-vous ? [Erreurs provenant de displayDialogAsync](#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-online"></a>Tirer parti d?une option de performances dans Office Online

La propri?t? `displayInIframe` est une propri?t? suppl?mentaire dans l?objet de configuration que vous pouvez transmettre ? `displayDialogAsync`. Lorsque cette propri?t? est d?finie sur `true` et que le compl?ment est en cours d?ex?cution dans un document ouvert dans Office Online, la bo?te de dialogue s?ouvre sous la forme d?un iFrame flottant et non d?une fen?tre ind?pendante. Elle s?ouvre ainsi plus rapidement. Voici un exemple :

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

La valeur par d?faut est `false`, ce qui revient ? omettre enti?rement la propri?t?. Si le compl?ment n?est pas ex?cut? dans Office Online, le `displayInIframe` est ignor?.

> [!NOTE]
> Vous ne devez **pas** utiliser `displayInIframe: true` si la bo?te de dialogue redirige ? un moment donn? l?utilisateur vers une page qui ne peut pas ?tre ouverte dans un iFrame. Par exemple, les pages de connexion de nombreux services web connus, comme un compte Microsoft et Google, ne peuvent pas ?tre ouvertes dans un iFrame.

### <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Envoi d?informations ? la page h?te ? partir de la bo?te de dialogue

La bo?te de dialogue ne peut pas communiquer avec la page h?te dans le volet Office, sauf si :

- la page active dans la bo?te de dialogue se trouve dans le m?me domaine que la page h?te ;
- la biblioth?que JavaScript Office est charg?e dans la page. (Comme n?importe quelle page qui utilise la biblioth?que JavaScript Office, le script de la page doit attribuer une m?thode ? la propri?t? `Office.initialize`, bien qu?il puisse s?agir d?une m?thode vide. Pour plus d?informations, voir [Initialisation de votre compl?ment](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)

Le code de la page de bo?te de dialogue utilise la fonction `messageParent` pour envoyer une valeur bool?enne ou un message de type cha?ne ? la page h?te. La cha?ne peut ?tre un mot, une phrase, un blob XML, un JSON converti en cha?ne ou un autre ?l?ment pouvant ?tre s?rialis? en cha?ne. Voici un exemple :

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - La fonction `messageParent` est l?une des deux *seules* API Office pouvant ?tre appel?es dans la bo?te de dialogue. L?autre est `Office.context.requirements.isSetSupported`. Pour plus d?informations, consultez la rubrique relative ? la [sp?cification d?h?tes Office et de conditions requises d?API](specify-office-hosts-and-api-requirements.md).
> - La fonction `messageParent` peut uniquement ?tre appel?e sur une page ayant le m?me domaine (y compris les m?mes protocole et port) que la page h?te.

Dans l?exemple suivant, `googleProfile` est une version convertie en cha?ne du profil Google de l?utilisateur.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

La page h?te doit ?tre configur?e de fa?on ? recevoir le message. Pour ce faire, ajoutez un param?tre de rappel ? l?appel d?origine de `displayDialogAsync`. Le rappel attribue un gestionnaire ? l??v?nement `DialogMessageReceived`. Voici un exemple :

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
> - Office transmet un objet [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) au rappel. Il repr?sente le r?sultat de la tentative d?ouverture de la bo?te de dialogue. Il ne repr?sente pas le r?sultat de tous les ?v?nements dans la bo?te de dialogue. Pour plus d?informations sur cette distinction, consultez la section [Gestion des erreurs et des ?v?nements](#handle-errors-and-events).
> - La propri?t? `value` de `asyncResult` est d?finie sur un objet [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog), qui existe dans la page h?te, pas dans le contexte d?ex?cution de la bo?te de dialogue.
> - est la fonction qui g?re l??v?nement. Vous pouvez lui donner le nom que vous souhaitez.`processMessage`
> - La variable `dialog` est d?clar?e avec une port?e plus large que le rappel, car elle est ?galement r?f?renc?e dans `processMessage`.

Voici un exemple simple de gestionnaire pour l??v?nement `DialogMessageReceived` :

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - Office transmet l?objet `arg` au gestionnaire. Sa propri?t? `message` est la valeur bool?enne ou la cha?ne envoy?e par l?appel de `messageParent` dans la bo?te de dialogue. Dans cet exemple, il s?agit d?une repr?sentation convertie en cha?ne du profil de l?utilisateur ? partir d?un service tel qu?un compte Microsoft ou Google, qui est donc d?s?rialis? en objet avec `JSON.parse`.
> - L?impl?mentation `showUserName` n?est pas visible. Elle peut afficher un message de bienvenue personnalis? dans le volet Office.

Lorsque l?intervention de l?utilisateur sur la bo?te de dialogue est termin?e, votre gestionnaire de messages doit fermer la bo?te de dialogue, comme indiqu? dans cet exemple.

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - L?objet `dialog` doit ?tre le m?me que celui renvoy? par l?appel de `displayDialogAsync`.
> - L?appel de `dialog.close` indique ? Office de fermer imm?diatement la bo?te de dialogue.

Pour voir un exemple de compl?ment qui utilise ces techniques, consultez la rubrique relative ? l?[exemple d?API de dialogue de compl?ment Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Si le compl?ment a besoin d?ouvrir une autre page du volet Office apr?s avoir re?u le message, vous pouvez utiliser la m?thode `window.location.replace` (ou `window.location.href`) en tant que derni?re ligne du gestionnaire. Voici un exemple :

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

Pour voir un exemple de compl?ment qui effectue ce type d?action, consultez la rubrique relative ? l?exemple [Ins?rer des graphiques Excel ? l?aide de Microsoft Graph dans un compl?ment PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

#### <a name="conditional-messaging"></a>Messagerie conditionnelle
?tant donn? que vous pouvez envoyer plusieurs appels `messageParent` ? partir de la bo?te de dialogue, mais que vous n?avez qu?un seul gestionnaire dans la page h?te pour l??v?nement `DialogMessageReceived`, le gestionnaire doit utiliser la logique conditionnelle pour distinguer les diff?rents messages. Par exemple, si la bo?te de dialogue invite un utilisateur ? se connecter ? un fournisseur d?identit? tel qu?un compte Microsoft ou Google, elle envoie le profil de l?utilisateur sous la forme d?un message. Si l?authentification ?choue, la bo?te de dialogue envoie des informations sur l?erreur ? la page h?te, comme dans l?exemple suivant :

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
> - La variable `loginSuccess` serait initialis?e en lisant la r?ponse HTTP ? partir du fournisseur d?identit?.
> - L?impl?mentation des fonctions `getProfile` et `getError` n?est pas affich?e. Chacune obtient des donn?es ? partir d?un param?tre de requ?te ou du corps de la r?ponse HTTP.
> - Des objets anonymes de diff?rents types sont envoy?s selon que la connexion a r?ussi ou non. Tous deux ont une propri?t? `messageType`, mais un a une propri?t? `profile` et l?autre une propri?t? `error`.

Pour obtenir des exemples qui utilisent la messagerie conditionnelle, consultez les rubriques suivantes :
- [Compl?ment Office qui utilise le service Auth0 pour simplifier la connexion sociale](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Compl?ment Office qui utilise le service OAuth.io pour simplifier l?acc?s aux services en ligne populaires](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

Le code du gestionnaire dans la page h?te utilise la valeur de la propri?t? `messageType` pour cr?er une branche comme le montre l?exemple suivant. Notez que la fonction `showUserName` est identique ? celle de l?exemple pr?c?dent et que la fonction `showNotification` affiche l?erreur dans l?interface utilisateur de la page h?te.

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

### <a name="closing-the-dialog-box"></a>Fermeture de la bo?te de dialogue

Vous pouvez impl?menter un bouton de fermeture dans la bo?te de dialogue. Pour ce faire, le gestionnaire d??v?nements Click du bouton doit utiliser `messageParent` pour indiquer ? la page h?te que vous avez cliqu? sur le bouton. Voici un exemple :

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

Le gestionnaire de la page h?te pour `DialogMessageReceived` appelle `dialog.close`, comme dans cet exemple. (Consultez les exemples pr?c?dents qui montrent comment l?objet Dialog est initialis?.)


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

Pour voir un exemple qui utilise cette technique, consultez le [mod?le de conception de navigation de bo?te de dialogue](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation) dans le r?f?rentiel de [mod?les de conception de l?exp?rience utilisateur pour compl?ments Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

M?me lorsque vous ne disposez pas de votre propre IU de fermeture de bo?te de dialogue, un utilisateur final peut fermer la bo?te de dialogue en choisissant le **X** dans le coin sup?rieur droit. Cette action d?clenche l??v?nement `DialogEventReceived`. Si votre volet h?te a besoin de savoir quand cela se produit, il doit d?clarer un gestionnaire pour cet ?v?nement. Pour plus d?informations, consultez la section [Erreurs et ?v?nements dans la fen?tre de dialogue](#errors-and-events-in-the-dialog-window).

## <a name="handle-errors-and-events"></a>Gestion des erreurs et des ?v?nements

Votre code doit g?rer deux cat?gories d??v?nements :

- les erreurs renvoy?es par l?appel de `displayDialogAsync` car la bo?te de dialogue ne peut pas ?tre cr??e ;
- les erreurs, et autres ?v?nements, dans la fen?tre de dialogue.

### <a name="errors-from-displaydialogasync"></a>Erreurs provenant de displayDialogAsync

En plus des erreurs syst?me et de plateforme g?n?rales, trois erreurs sont propres ? l?appel de `displayDialogAsync`.

|Num?ro de code|Signification|
|:-----|:-----|
|12004|Le domaine de l?URL transmis ? `displayDialogAsync` n?est pas approuv?. Le domaine doit ?tre le m?me domaine que celui de la page h?te (y compris le protocole et le num?ro de port).|
|12005|L?URL transmise ? `displayDialogAsync` utilise le protocole HTTP. C?est le protocole HTTPS qui est requis. (Dans certaines versions d?Office, le message d?erreur renvoy? avec le code 12005 est identique ? celui renvoy? avec le code 12004.)|
|<span id="12007">12007</span>|Une bo?te de dialogue est d?j? ouverte ? partir de cette fen?tre h?te. Une fen?tre h?te, par exemple un volet Office, ne peut avoir qu?une seule bo?te de dialogue ouverte ? la fois.|

Lorsque `displayDialogAsync` est appel?, il transmet toujours un objet [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) ? sa fonction de rappel. Lorsque l?appel r?ussit (autrement dit, que la fen?tre de dialogue est ouverte), la propri?t? `value` de l?objet `AsyncResult` est un objet [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog). Vous trouverez un exemple dans la section [Envoi d?informations ? la page h?te ? partir de la bo?te de dialogue](#send-information-from-the-dialog-box-to-the-host-page). Lorsque l?appel de `displayDialogAsync` ?choue, la fen?tre n?est pas cr??e, la propri?t? `status` de l?objet `AsyncResult` est d?finie sur ? failed ? et la propri?t? `error` de l?objet est renseign?e. Vous devez toujours disposer d?un rappel qui teste `status` et r?pond lorsqu?il s?agit d?une erreur. Voici un exemple de code qui signale simplement le message d?erreur, quel que soit son num?ro de code :

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === "failed") {
        showNotification(asynceResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

### <a name="errors-and-events-in-the-dialog-window"></a>Erreurs et ?v?nements dans la fen?tre de dialogue

Trois erreurs et ?v?nements, d?sign?s par leur num?ro de code, dans la bo?te de dialogue d?clencheront un ?v?nement `DialogEventReceived` dans la page h?te.

|Num?ro de code|Signification|
|:-----|:-----|
|12002|Un des ?l?ments suivants :<br> - Aucune page n?existe ? l?URL qui a ?t? transmise ? `displayDialogAsync`.<br> - La page qui a ?t? transmise ? `displayDialogAsync` a ?t? charg?e, mais la bo?te de dialogue a ?t? redirig?e vers une page introuvable ou impossible ? charger, ou a ?t? redirig?e vers une URL dont la syntaxe n?est pas valide.|
|12003|La bo?te de dialogue a ?t? redirig?e vers une URL avec le protocole HTTP. C?est le protocole HTTPS qui est requis.|
|12006|La bo?te de dialogue a ?t? ferm?e, g?n?ralement parce que l?utilisateur choisit le bouton **X**.|

Votre code peut attribuer un gestionnaire pour l??v?nement `DialogEventReceived` dans l?appel de `displayDialogAsync`. Voici un exemple simple :

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

Pour voir un exemple de gestionnaire pour l??v?nement `DialogEventReceived` qui cr?e des messages d?erreur personnalis?s pour chaque code d?erreur, consultez l?exemple suivant :

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

Pour voir un exemple de compl?ment qui g?re les erreurs de cette fa?on, consultez la rubrique relative ? l?[exemple d?API de dialogue de compl?ment Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).


## <a name="pass-information-to-the-dialog-box"></a>Transmission d?informations ? la bo?te de dialogue

Parfois, la page h?te doit transmettre des informations ? la bo?te de dialogue. Pour ce faire, il existe deux moyens :

- ajouter des param?tres de requ?te ? l?URL qui est transmise ? `displayDialogAsync` ;
- stocker les informations ? un emplacement auquel ? la fois la fen?tre h?te et la bo?te de dialogue ont acc?s. Les deux fen?tres ne partagent pas un stockage de session commun, mais *si elles ont le m?me domaine* (y compris le m?me num?ro de port, le cas ?ch?ant), elles utilisent un [stockage local](http://www.w3schools.com/html/html5_webstorage.asp) commun.

### <a name="use-local-storage"></a>Utilisation du stockage local

Pour utiliser le stockage local, votre code appelle la m?thode `setItem` de l?objet `window.localStorage` dans la page h?te avant l?appel de `displayDialogAsync`, comme dans l?exemple suivant :

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Le code dans la fen?tre de dialogue lit l??l?ment lorsqu?il est n?cessaire, comme dans l?exemple suivant :

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

Pour obtenir des exemples de compl?ments qui utilisent le stockage local de cette fa?on, consultez les rubriques suivantes :

- [Compl?ment Office qui utilise le service Auth0 pour simplifier la connexion sociale](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Compl?ment Office qui utilise le service OAuth.io pour simplifier l?acc?s aux services en ligne populaires](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

### <a name="use-query-parameters"></a>Utiliser les param?tres de requ?te

L?exemple suivant montre comment transmettre des donn?es ? l?aide d?un param?tre de requ?te :

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Pour voir un exemple qui utilise cette technique, consultez la rubrique relative ? l?exemple [Ins?rer des graphiques Excel ? l?aide de Microsoft Graph dans un compl?ment PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Le code dans votre fen?tre de dialogue peut analyser l?URL et lire la valeur du param?tre.

> [!NOTE]
> Office ajoute automatiquement un param?tre de requ?te appel? `_host_info` ? l?URL qui est transmise ? `displayDialogAsync`. (Il est ajout? apr?s vos param?tres de requ?te personnalis?s, le cas ?ch?ant. Il n?est pas ajout? ? toutes les autres URL auxquelles la bo?te de dialogue acc?de.) Microsoft peut modifier le contenu de cette valeur, ou le supprimer enti?rement, ? l?avenir, donc votre code ne doit pas le lire. La m?me valeur est ajout?e au stockage de session de la bo?te de dialogue. L? encore, *votre code ne doit ni lire, ni ?crire cette valeur*.

## <a name="use-the-dialog-apis-to-show-a-video"></a>Utilisation des API de dialogue pour afficher une vid?o

Pour afficher une vid?o dans une bo?te de dialogue :

1.  Cr?ez une page dont seul le contenu est un iFrame. L?attribut `src` de l?iFrame pointe vers une vid?o en ligne. Le protocole de l?URL de la vid?o doit ?tre HTTP**S**. Dans cet article, nous appellerons cette page ? video.dialogbox.html ?. Voici un exemple de marques de r?vision :

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  La page video.dialogbox.html doit se trouver dans le m?me domaine que la page h?te.
3.  Utilisez un appel de `displayDialogAsync` dans la page h?te pour ouvrir video.dialogbox.html.
4.  Si votre compl?ment a besoin de savoir quand l?utilisateur ferme la bo?te de dialogue, inscrivez un gestionnaire pour l??v?nement `DialogEventReceived` et g?rez l??v?nement 12006. Pour plus d?informations, consultez la section [Erreurs et ?v?nements dans la fen?tre de dialogue](#errors-and-events-in-the-dialog-window).

Pour voir un exemple qui affiche une vid?o dans une bo?te de dialogue, consultez le [mod?le de conception de maquette de vid?o](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat) dans le r?f?rentiel de [mod?les de conception de l?exp?rience utilisateur pour compl?ments Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

![Capture d??cran d?une vid?o s?affichant dans une bo?te de dialogue de compl?ment](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a>Utilisation des API de dialogue dans un flux d?authentification

Le sc?nario principal des API de dialogue consiste ? activer l?authentification aupr?s d?un fournisseur de ressources ou d?identit? qui n?autorise pas l?ouverture de sa page de connexion dans un iframe, comme un compte Microsoft, Office 365, Google et Facebook.

> [!NOTE]
> Lorsque vous utilisez les API de dialogue pour ce sc?nario, n?utilisez *pas* l?option `displayInIframe: true` dans l?appel de `displayDialogAsync`. Reportez-vous ? la section [Tirer parti d?une option de performances dans Office Online](#take-advantage-of-a-performance-option-in-office-online) pr?c?demment dans cet article pour plus d?informations sur cette option.

Voici un flux d?authentification simple et standard :

1. La premi?re page qui s?ouvre dans la bo?te de dialogue est une page locale (ou toute autre ressource) qui est h?berg?e dans le domaine du compl?ment. Autrement dit, le domaine de la fen?tre h?te. Cette page peut avoir une IU simple indiquant ? Veuillez patienter, nous allons vous rediriger vers la page sur laquelle vous pouvez vous connecter ? *NOM DU FOURNISSEUR* ?. Le code dans cette page construit l?URL de la page de connexion du fournisseur d?identit? en utilisant les informations transmises ? la bo?te de dialogue, comme d?crit dans [Transmission d?informations ? la bo?te de dialogue](#pass-information-to-the-dialog-box).
2. La fen?tre de dialogue redirige alors l?utilisateur vers la page de connexion. L?URL inclut un param?tre de requ?te qui indique au fournisseur d?identit? de rediriger la fen?tre de dialogue une fois que l?utilisateur s?est connect? ? une page sp?cifique. Dans cet article, nous appellerons cette page ? redirectPage.html ?. (*Il doit s?agir d?une page ayant le m?me domaine que la fen?tre h?te*, car le seul moyen pour que la fen?tre de dialogue transmette les r?sultats de la tentative de connexion est un appel de `messageParent`, qui ne peut ?tre appel? que sur une page ayant le m?me domaine que la fen?tre h?te.)
2. Le service du fournisseur d?identit? traite la requ?te GET entrante ? partir de la fen?tre de dialogue. Si l?utilisateur est d?j? connect?, il redirige imm?diatement la fen?tre vers redirectPage.html et inclut les donn?es utilisateur sous la forme d?un param?tre de requ?te. Si l?utilisateur n?est pas encore connect?, la page de connexion du fournisseur appara?t dans la fen?tre et l?utilisateur se connecte. Pour la plupart des fournisseurs, si l?utilisateur ne parvient pas ? se connecter, le fournisseur affiche une page d?erreur dans la fen?tre de dialogue et ne redirige pas vers redirectPage.html. L?utilisateur doit fermer la fen?tre en s?lectionnant le **X** dans le coin. Si l?utilisateur se connecte avec succ?s, la fen?tre de dialogue est redirig?e vers redirectPage.html et les donn?es utilisateur sont incluses sous la forme d?un param?tre de requ?te.
3. Lorsque la page redirectPage.html s?ouvre, elle appelle `messageParent` pour indiquer le succ?s ou l??chec ? la page h?te et ?ventuellement indiquer ?galement des donn?es utilisateur ou des donn?es d?erreur.
4. L??v?nement `DialogMessageReceived` se d?clenche dans la page h?te, et son gestionnaire ferme la fen?tre de dialogue et effectue ?ventuellement d?autres traitements du message.

Pour voir des exemples de compl?ments qui utilisent ce mod?le, consultez les pages suivantes :

- [Ins?rer des graphiques Excel ? l?aide de Microsoft Graph dans un compl?ment PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) : La ressource qui s?ouvre initialement dans la fen?tre de la bo?te de dialogue est une m?thode du contr?leur qui ne dispose d?aucun affichage propre. Elle redirige l?utilisateur vers la page de connexion Office 365.
- [Authentification client Office 365 du compl?ment Office pour AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth) : La ressource qui s?ouvre initialement dans la fen?tre de dialogue est une page.

#### <a name="support-multiple-identity-providers"></a>Prise en charge de plusieurs fournisseurs d?identit?

Si votre compl?ment offre ? l?utilisateur le choix entre plusieurs fournisseurs, tels qu?un compte Microsoft, Google ou Facebook, vous avez besoin d?une premi?re page locale (voir section pr?c?dente) qui fournit une IU permettant ? l?utilisateur de s?lectionner un fournisseur. La s?lection d?clenche la construction de l?URL de connexion et la redirection vers celle-ci.

Pour voir un exemple qui utilise ce mod?le, consultez la rubrique relative ? l?exemple [Compl?ment Office qui utilise le service Auth0 pour simplifier la connexion aux r?seaux sociaux](https://github.com/OfficeDev/Office-Add-in-Auth0).

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a>Autorisation du compl?ment pour une ressource externe

Sur le web nouvelle g?n?ration, les applications web sont des principaux de s?curit? au m?me titre que les utilisateurs, et l?application a sa propre identit? et ses propres autorisations pour une ressource en ligne comme Office 365, Google Plus, Facebook ou LinkedIn. L?application est inscrite aupr?s du fournisseur de ressources avant d??tre d?ploy?e. L?inscription inclut :

- la liste des autorisations dont l?application a besoin pour les ressources d?un utilisateur ;
- l?URL ? laquelle le service de ressources doit renvoyer un jeton d?acc?s lorsque l?application acc?de au service.  

Lorsqu?un utilisateur appelle une fonction dans l?application qui acc?de aux donn?es de l?utilisateur dans le service de ressources, l?utilisateur est invit? ? se connecter au service, puis ? accorder ? l?application les autorisations dont elle a besoin pour les ressources de l?utilisateur. Ensuite, le service redirige la fen?tre de connexion vers l?URL pr?c?demment inscrite et transmet le jeton d?acc?s. L?application utilise le jeton d?acc?s pour acc?der aux ressources de l?utilisateur.

Vous pouvez utiliser les API de dialogue pour g?rer ce processus ? l?aide d?un flux semblable ? celui d?crit pour la connexion des utilisateurs. Les seules diff?rences sont les suivantes :

- Si l?utilisateur n?a pas pr?alablement accord? ? l?application les autorisations n?cessaires, il est invit? ? le faire dans la bo?te de dialogue apr?s la connexion.
- La fen?tre de dialogue envoie le jeton d?acc?s ? la fen?tre h?te en utilisant `messageParent` pour envoyer le jeton d?acc?s converti en cha?ne ou en stockant jeton d?acc?s ? un emplacement o? la fen?tre h?te peut le r?cup?rer. Le jeton a une limite de temps, mais tant qu?elle n?est pas ?coul?e, la fen?tre h?te peut l?utiliser pour acc?der directement aux ressources de l?utilisateur sans demander d?autre confirmation.

Les exemples suivants utilisent les API de dialogue ? cet effet :
- [Ins?rer des graphiques Excel ? l?aide de Microsoft Graph dans un compl?ment PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) : stocke le jeton d?acc?s dans une base de donn?es.
- [Compl?ment Office qui utilise le service OAuth.io pour simplifier l?acc?s aux services en ligne populaires](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

Pour plus d?informations sur l?authentification et l?autorisation dans des compl?ments, consultez les rubriques suivantes :
- [Autoriser des services externes dans votre compl?ment Office](auth-external-add-ins.md)
- [Biblioth?que d?applications d?assistance des API JavaScript Office](https://github.com/OfficeDev/office-js-helpers)


## <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Utilisation de l?API de dialogue Office avec des applications ? page unique et routage c?t? client

Si votre compl?ment utilise le routage c?t? client, comme le font les applications ? page unique en r?gle g?n?rale, vous avez la possibilit? de transmettre l?URL d?un itin?raire ? la m?thode [displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync), au lieu de l?URL de la page HTML compl?te et distincte.

> [!IMPORTANT]
>La bo?te de dialogue se trouve dans une nouvelle fen?tre avec son propre contexte d?ex?cution. Si vous transmettez un itin?raire, votre page de base et son code d?initialisation et d?amor?age s?ex?cutent ? nouveau dans ce nouveau contexte, et toutes les variables sont d?finies sur leurs valeurs initiales dans la fen?tre de dialogue. Par cons?quent, cette technique lance une deuxi?me instance de votre application dans la fen?tre de dialogue. Le code qui modifie des variables dans la fen?tre de dialogue ne change pas la version du volet Office des m?mes variables. De m?me, la fen?tre de dialogue poss?de son propre stockage de session, qui n?est pas accessible ? partir du code dans le volet Office.
