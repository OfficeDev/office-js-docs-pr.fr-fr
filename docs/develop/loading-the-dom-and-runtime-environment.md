---
title: Chargement du DOM et de l’environnement d’exécution
description: ''
ms.date: 07/01/2019
localization_priority: Normal
ms.openlocfilehash: 06b3e9d65d29b257d34d2f4bdad81f464056e558
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325121"
---
# <a name="loading-the-dom-and-runtime-environment"></a>Chargement du DOM et de l’environnement d’exécution

Un complément doit s’assurer que le DOM et l’environnement d’exécution des Compléments Office ont été chargés avant d’exécuter sa propre logique personnalisée.

## <a name="startup-of-a-content-or-task-pane-add-in"></a>Démarrage d’un complément de contenu ou du volet Office

La figure suivante illustre le flux des événements impliqués au démarrage d’un complément de contenu ou du volet Office dans Excel, PowerPoint, Project ou Word.

![Flux des événements au démarrage d’un complément de contenu ou du volet Office](../images/office15-app-sdk-loading-dom-agave-runtime.png)

Les événements suivants se produisent lors du démarrage d’un complément de contenu ou du volet Office :

1. L’utilisateur ouvre un document qui contient déjà un complément ou insère un complément dans le document.

2. L’application hôte Office lit le manifeste XML du complément à partir d’AppSource, d’un catalogue d’applications sur SharePoint ou du catalogue de dossiers partagés duquel il provient.

3. L’application hôte Office ouvre la page HTML du complément dans un contrôle de navigateur.

    Les deux étapes suivantes, 4 et 5, se produisent de manière asynchrone et parallèlement. C’est pour cela que le code de votre complément doit veiller à ce que le chargement du DOM et de l’environnement d’exécution du complément soit terminé avant de continuer.

4. Le contrôle de navigateur charge le DOM et le corps HTML, et appelle le gestionnaire d' `window.onload` événements pour l’événement.

5. L’application hôte Office charge l’environnement d’exécution, lequel télécharge et met en cache l’API JavaScript pour les fichiers de bibliothèque JavaScript à partir du serveur de réseau de distribution de contenu, puis appelle le gestionnaire d’événements du complément pour l’événement [initialize](/javascript/api/office#office-initialize-reason-) de l’objet [Office](/javascript/api/office) si un gestionnaire lui a été affecté. Il vérifie alors également si des rappels (ou des fonctions `then()` chaînées) ont été transmis (ou chaînées) au gestionnaire `Office.onReady`. Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady`, voir [Initialisation de votre complément](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).

6. Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.


## <a name="startup-of-an-outlook-add-in"></a>Démarrage d’un complément Outlook

La figure suivante illustre le flux des événements impliqués au démarrage d’un complément Outlook exécuté sur un ordinateur de bureau, une tablette ou un smartphone.

![Flux des événements au démarrage du complément Outlook](../images/outlook15-loading-dom-agave-runtime.png)

Les événements suivants se produisent lors du démarrage d’un complément Outlook :

1. Lorsqu’Outlook démarre, il lit les manifestes XML pour les compléments Outlook qui ont été installés pour le compte de messagerie de l’utilisateur.

2. L’utilisateur sélectionne un élément dans Outlook.

3. Si l’élément sélectionné répond aux conditions d’activation d’un complément Outlook, Outlook active le complément et affiche son bouton dans l’interface utilisateur.

4. Si l’utilisateur clique sur le bouton pour démarrer le complément Outlook, Outlook ouvre la page HTML dans un contrôle de navigateur. Les deux étapes suivantes, 5 et 6, se produisent en parallèle.

5. Le contrôle de navigateur charge le DOM et le corps HTML, et appelle le gestionnaire d' `onload` événements pour l’événement.

6. Outlook charge l’environnement d’exécution, lequel télécharge et met en cache l’API JavaScript pour les fichiers de bibliothèque JavaScript à partir du serveur de réseau de distribution de contenu, puis appelle le gestionnaire d’événements du complément pour l’événement [initialize](/javascript/api/office#office-initialize-reason-) de l’objet [Office](/javascript/api/office) du complément si un gestionnaire lui a été affecté. Il vérifie alors également si des rappels (ou des fonctions `then()` chaînées) ont été transmis (ou chaînées) au gestionnaire `Office.onReady`. Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady`, voir [Initialisation de votre complément](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).

7. Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.


## <a name="checking-the-load-status"></a>Vérification du statut de chargement

Vous pouvez vérifier que le chargement du DOM et de l’environnement d’exécution est bien terminé en utilisant la fonction jQuery [.ready()](https://api.jquery.com/ready/) : `$(document).ready()`. Par exemple, le gestionnaire `onReady` d’événements suivant vérifie que le DOM est chargé pour la première fois avant l’exécution du code spécifique à l’initialisation du complément. Par la suite `onReady` , le gestionnaire continue d’utiliser la propriété [Mailbox. Item](/javascript/api/outlook/office.mailbox) pour obtenir l’élément actuellement sélectionné dans Outlook et appelle la fonction principale du complément, `initDialer`.

```js
Office.onReady()
    .then(
        // Checks for the DOM to load.
        $(document).ready(function () {
            // After the DOM is loaded, add-in-specific code can run.
            var mailbox = Office.context.mailbox;
            _Item = mailbox.item;
            initDialer();
        });
);
```

Vous pouvez également utiliser le même code dans un `initialize` gestionnaire d’événements comme illustré dans l’exemple suivant.

```js
Office.initialize = function () {
    // Checks for the DOM to load.
    $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        var mailbox = Office.context.mailbox;
        _Item = mailbox.item;
        initDialer();
    });
}
```

Cette même technique peut être utilisée dans les `onReady` gestionnaires `initialize` ou des compléments Office.

Le numéroteur téléphonique fourni comme exemple de complément Outlook présente une approche légèrement différente, puisqu’il utilise uniquement JavaScript pour vérifier ces mêmes conditions. 

> [!IMPORTANT]
> Même si aucune tâche d’initialisation n’est à effectuer dans votre complément, vous devez inclure au moins un appel de `Office.onReady` la fonction de `Office.initialize` gestionnaire d’événements minimal, comme illustré dans les exemples suivants.
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> Si vous n’appelez `Office.onReady` pas ou n’assignez pas de gestionnaire d' `Office.initialize` événements, votre complément peut déclencher une erreur lors de son démarrage. En outre, si un utilisateur essaie d’utiliser votre complément avec un client web Office, notamment Excel, PowerPoint ou Outlook, l’exécution du complément échouera.
>
> Si votre complément comprend plusieurs pages, chaque fois qu’il charge une nouvelle page, celle-ci doit appeler `Office.onReady` ou assigner `Office.initialize` un gestionnaire d’événements.

## <a name="see-also"></a>Voir aussi

- [Présentation de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [Initialiser votre complément Office](initialize-add-in.md)