---
title: Transmission de données et de messages à une boîte de dialogue à partir de sa page hôte
description: Découvrez comment transmettre des données à une boîte de dialogue à partir de la page hôte à l’aide des API messageChild et DialogParentMessageReceived.
ms.date: 04/16/2020
localization_priority: Normal
ms.openlocfilehash: 3bef98294b15c2787b707cee4861cc9932f98166
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609407"
---
# <a name="passing-data-and-messages-to-a-dialog-box-from-its-host-page-preview"></a>Transmission de données et de messages à une boîte de dialogue à partir de sa page hôte (aperçu)

Votre complément peut envoyer des messages à partir de la [page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) vers une boîte de dialogue à l’aide de la méthode [messageChild](/javascript/api/office/office.dialog#messagechild-message-) de l’objet [Dialog](/javascript/api/office/office.dialog) .

> [!Important]
>
> - Les API décrites dans cet article sont en aperçu. Elles sont disponibles pour les développeurs dans le cas d’expérimentation ; mais ne doit pas être utilisé dans un complément de production. Tant que cette API n’est pas publiée, utilisez les techniques décrites dans [transmettre les informations à la boîte de dialogue](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) des compléments de production.
> - Les API décrites dans cet article nécessitent Office 365 (la version avec abonnement d’Office). Vous devez utiliser la version et le build mensuels les plus récents du canal du programme Insider. Vous devez participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://insider.office.com). Veuillez noter que lorsqu’une build est basée sur le canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu est désactivée pour cette version.
> - Dans l’étape initiale de l’aperçu, les API sont prises en charge dans Excel, PowerPoint et Word ; mais pas dans Outlook.
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="use-messagechild-from-the-host-page"></a>Utiliser `messageChild()` à partir de la page hôte

Lorsque vous appelez l’API de boîte de dialogue Office pour ouvrir une boîte de dialogue, un objet [Dialog](/javascript/api/office/office.dialog) est renvoyé. Elle doit être assignée à une variable, qui a généralement une portée plus élevée que la méthode [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) , car l’objet sera référencé par d’autres méthodes. Voici un exemple :

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

Cet `Dialog` objet est doté d’une méthode [messageChild](/javascript/api/office/office.dialog#messagechild-message-) qui envoie une chaîne ou des données JSON à la boîte de dialogue. Cela déclenche un `DialogParentMessageReceived` événement dans la boîte de dialogue. Votre code doit gérer cet événement, comme indiqué dans la section suivante.

Imaginez un scénario dans lequel l’interface utilisateur de la boîte de dialogue doit correspondre à la feuille de calcul active et la position de cette feuille de calcul par rapport aux autres feuilles de calcul. Dans l’exemple suivant, `sheetPropertiesChanged` envoie les propriétés de feuille de calcul Excel dans la boîte de dialogue. Dans ce cas, la feuille de calcul active est nommée « ma feuille » et est la seconde feuille du classeur. Les données sont encapsulées dans un objet qui est JSON afin de pouvoir être transmis à `messageChild` .

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Gérer DialogParentMessageReceived dans la boîte de dialogue

Dans le JavaScript de la boîte de dialogue, inscrivez un gestionnaire pour l' `DialogParentMessageReceived` événement à l’aide de la méthode [UI. addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) . Cette opération s’effectue généralement dans les [méthodes Office. onReady ou Office. Initialize](initialize-add-in.md). Voici un exemple :

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Ensuite, définissez le `onMessageFromParent` Gestionnaire. Le code suivant poursuit l’exemple de la section précédente. Notez qu’Office transmet un argument au gestionnaire et que la `message` propriété de l’objet argument contient la chaîne de la page hôte. Dans cet exemple, le message est reconverti en objet et jQuery est utilisé pour définir le titre supérieur de la boîte de dialogue de sorte qu’il corresponde au nouveau nom de la feuille de calcul.

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

Il est recommandé de vérifier que votre gestionnaire est correctement enregistré. Pour ce faire, vous pouvez transmettre un rappel à la `addHandlerAsync` méthode qui s’exécute lorsque la tentative d’enregistrement du gestionnaire est terminée. Utilisez le gestionnaire pour consigner ou afficher une erreur si le gestionnaire n’a pas été enregistré correctement. Voici un exemple. Notez qu' `reportError` il s’agit d’une fonction, non définie ici, qui enregistre ou affiche l’erreur.

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

## <a name="conditional-messaging"></a>Messagerie conditionnelle

Étant donné que vous pouvez effectuer plusieurs `messageChild` appels à partir de la page hôte, mais que vous n’avez qu’un seul gestionnaire dans la boîte de dialogue de l' `DialogParentMessageReceived` événement, le gestionnaire doit utiliser une logique conditionnelle pour distinguer les différents messages. Vous pouvez effectuer cette opération d’une manière parfaitement parallèle à la façon dont vous structurez la messagerie conditionnelle lorsque la boîte de dialogue envoie un message à la page hôte, comme décrit dans la section [messagerie conditionnelle](dialog-api-in-office-add-ins.md#conditional-messaging).
