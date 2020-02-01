---
title: Gestion des erreurs et des événements dans la boîte de dialogue Office
description: Indique comment intercepter et gérer les erreurs lors de l’ouverture et de l’utilisation de la boîte de dialogue Office
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: a35131a46dc9f5edc18df37495abe5d8c2c5ad2a
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/01/2020
ms.locfileid: "41650085"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a>Gestion des erreurs et des événements dans la boîte de dialogue Office

Cet article explique comment intercepter et gérer les erreurs lors de l’ouverture de la boîte de dialogue et des erreurs qui se produisent dans la boîte de dialogue.

> [!NOTE]
> Cet article suppose que vous êtes familiarisé avec les notions de base de l’utilisation de l’API de boîte de dialogue Office, comme décrit dans la rubrique [use the Office Dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).
> 
> Voir aussi [meilleures pratiques et règles pour l’API de boîte de dialogue Office](dialog-best-practices.md).

Votre code doit gérer deux catégories d’événements :

- les erreurs renvoyées par l’appel de `displayDialogAsync` car la boîte de dialogue ne peut pas être créée ;
- Erreurs et autres événements, dans la boîte de dialogue.

## <a name="errors-from-displaydialogasync"></a>Erreurs provenant de displayDialogAsync

En plus des erreurs système et de plateforme générales, quatre erreurs sont propres à `displayDialogAsync`l’appel.

|Numéro de code|Signification|
|:-----|:-----|
|12004|Le domaine de l’URL transmis à `displayDialogAsync` n’est pas approuvé. Le domaine doit être le même domaine que celui de la page hôte (y compris le protocole et le numéro de port).|
|12005|L’URL transmise à `displayDialogAsync` utilise le protocole HTTP. C’est le protocole HTTPS qui est requis. (Dans certaines versions d’Office, le texte du message d’erreur renvoyé avec 12005 est le même que celui renvoyé pour 12004.)|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|Une boîte de dialogue est déjà ouverte à partir de cette fenêtre hôte. Une fenêtre hôte, par exemple un volet Office, ne peut avoir qu’une seule boîte de dialogue ouverte à la fois.|
|12009|L’utilisateur a choisi d’ignorer la boîte de dialogue. Cette erreur peut se produire dans Office sur le Web, où les utilisateurs peuvent choisir de ne pas autoriser un complément à présenter une boîte de dialogue. Pour plus d’informations, consultez [la rubrique gestion des bloqueurs de fenêtres publicitaires intempestives avec Office sur le Web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).|

Lorsque `displayDialogAsync` est appelé, il transmet un objet [asyncResult](/javascript/api/office/office.asyncresult) à sa fonction de rappel. Une fois l’appel réussi, la boîte de dialogue est ouverte et la `value` propriété de l' `AsyncResult` objet est un objet [Dialog](/javascript/api/office/office.dialog) . Pour obtenir un exemple, reportez-vous [à la rubrique envoyer des informations de la boîte de dialogue à la page hôte](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page). Lorsque l’appel à `displayDialogAsync` échoue, la boîte de dialogue n’est pas créée `status` , la propriété `AsyncResult` de l’objet est `Office.AsyncResultStatus.Failed`définie sur et `error` la propriété de l’objet est remplie. Vous devez toujours fournir un rappel qui teste `status` le et répond lorsqu’il s’agit d’une erreur. Pour obtenir un exemple qui signale le message d’erreur quel que soit son numéro de code, consultez le code suivant. (La `showNotification` fonction, non définie dans cet article, affiche ou consigne l’erreur. Pour obtenir un exemple de la façon dont vous pouvez implémenter cette fonction dans votre complément, consultez la rubrique [exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

## <a name="errors-and-events-in-the-dialog-box"></a>Erreurs et événements dans la boîte de dialogue

Trois erreurs et événements dans la boîte de dialogue déclencheront `DialogEventReceived` un événement dans la page hôte. Pour un rappel de ce qu’est une page hôte, voir [ouvrir une boîte de dialogue à partir d’une page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).

|Numéro de code|Signification|
|:-----|:-----|
|12002|Un des éléments suivants :<br> - Aucune page n’existe à l’URL qui a été transmise à `displayDialogAsync`.<br> -La page qui a été transmise au `displayDialogAsync` chargement, mais la boîte de dialogue a été redirigée vers une page qu’elle ne peut pas trouver ou chargée, ou elle a été dirigée vers une URL avec une syntaxe incorrecte.|
|12003|La boîte de dialogue a été redirigée vers une URL avec le protocole HTTP. C’est le protocole HTTPS qui est requis.|
|12006|La boîte de dialogue a été fermée, généralement parce que l’utilisateur a cliqué sur le bouton **Fermer** **X**.|

Votre code peut attribuer un gestionnaire pour l’événement `DialogEventReceived` dans l’appel de `displayDialogAsync`. Voici un exemple simple :

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

Pour voir un exemple de gestionnaire pour l’événement `DialogEventReceived` qui crée des messages d’erreur personnalisés pour chaque code d’erreur, consultez l’exemple suivant :

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

Pour voir un exemple de complément qui gère les erreurs de cette façon, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).
