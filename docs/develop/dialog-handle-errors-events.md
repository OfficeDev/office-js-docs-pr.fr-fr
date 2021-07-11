---
title: Gestion des erreurs et des événements dans la boîte de dialogue Office
description: Décrit comment éviter et gérer les erreurs lors de l’ouverture et de l’utilisation de la boîte Office dialogue
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: be1fb8bcd30b47ac6399657d928d3cad7f857f39
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349895"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a>Gestion des erreurs et des événements dans la boîte de dialogue Office

Cet article explique comment prendre en charge les erreurs lors de l’ouverture de la boîte de dialogue et les erreurs qui se produisent à l’intérieur de la boîte de dialogue.

> [!NOTE]
> Cet article présuppose que vous connaissez les principes de base de l’utilisation de l’API de boîte de dialogue Office, comme décrit dans [l’API](dialog-api-in-office-add-ins.md)de boîte de dialogue Office dans vos Office.
> 
> Voir aussi [Meilleures pratiques et règles pour l’API Office boîte de dialogue .](dialog-best-practices.md)

Votre code doit gérer deux catégories d’événements :

- les erreurs renvoyées par l’appel de `displayDialogAsync` car la boîte de dialogue ne peut pas être créée ;
- Erreurs et autres événements dans la boîte de dialogue.

## <a name="errors-from-displaydialogasync"></a>Erreurs provenant de displayDialogAsync

Outre les erreurs générales de plateforme et système, quatre erreurs sont spécifiques à `displayDialogAsync` l’appel.

|Numéro de code|Signification|
|:-----|:-----|
|12004|Le domaine de l’URL transmis à `displayDialogAsync` n’est pas approuvé. Le domaine doit être le même domaine que celui de la page hôte (y compris le protocole et le numéro de port).|
|12005|L’URL transmise à `displayDialogAsync` utilise le protocole HTTP. C’est le protocole HTTPS qui est requis. (Dans certaines versions de Office, le texte du message d’erreur renvoyé avec 12005 est identique à celui renvoyé pour 12004.)|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|Une boîte de dialogue est déjà ouverte à partir de cette fenêtre hôte. Une fenêtre hôte, par exemple un volet Office, ne peut avoir qu’une seule boîte de dialogue ouverte à la fois.|
|12009|L’utilisateur a choisi d’ignorer la boîte de dialogue. Cette erreur peut se produire dans Office sur le Web, où les utilisateurs peuvent choisir de ne pas autoriser un add-in à présenter une boîte de dialogue. Pour plus d’informations, voir [Gestion des bloqueurs de](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)fenêtres Office sur le Web .|

Lorsqu’elle est appelée, elle transmet un `displayDialogAsync` objet [AsyncResult](/javascript/api/office/office.asyncresult) à sa fonction de rappel. Lorsque l’appel réussit, la boîte de dialogue s’ouvre et la propriété de l’objet `value` `AsyncResult` est un objet [Dialog.](/javascript/api/office/office.dialog) Pour obtenir un exemple de cela, voir Envoyer des informations à partir de la boîte [de dialogue vers la page hôte.](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page) Lorsque l’appel échoue, la boîte de dialogue n’est pas créée, la propriété de l’objet est définie sur et la propriété de `displayDialogAsync` `status` `AsyncResult` l’objet est `Office.AsyncResultStatus.Failed` `error` remplie. Vous devez toujours fournir un rappel qui teste et répond en cas `status` d’erreur. Pour obtenir un exemple qui signale le message d’erreur, quel que soit son numéro de code, consultez le code suivant. (La `showNotification` fonction, non définie dans cet article, affiche ou enregistre l’erreur. Pour obtenir un exemple de la façon dont vous pouvez implémenter cette fonction dans votre application, voir Office’API de boîte de dialogue [de l’application.)](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)

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

Trois erreurs et événements dans la boîte de dialogue lèvent un `DialogEventReceived` événement dans la page hôte. Pour un rappel de ce qu’est une page hôte, voir Ouvrir une boîte de [dialogue à partir d’une page hôte.](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)

|Numéro de code|Signification|
|:-----|:-----|
|12002|Un des éléments suivants :<br> - Aucune page n’existe à l’URL qui a été transmise à `displayDialogAsync`.<br> - Page qui a été transmise au chargement, mais la boîte de dialogue a ensuite été redirigée vers une page qu’elle ne peut ni trouver ni charger, ou qui a été redirigée vers une URL dont la syntaxe n’est `displayDialogAsync` pas valide.|
|12003|La boîte de dialogue a été redirigée vers une URL avec le protocole HTTP. C’est le protocole HTTPS qui est requis.|
|12006|La boîte de dialogue a été fermée, généralement parce que l’utilisateur a choisi **le bouton** **Fermer X**.|

Votre code peut attribuer un gestionnaire pour l’événement `DialogEventReceived` dans l’appel de `displayDialogAsync`. Voici un exemple simple.

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

Pour obtenir un exemple de handler pour l’événement qui crée des messages d’erreur personnalisés pour chaque `DialogEventReceived` code d’erreur, voir l’exemple suivant.

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
