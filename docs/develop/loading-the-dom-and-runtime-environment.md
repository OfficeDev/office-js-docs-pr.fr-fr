---
title: Chargement du DOM et de l’environnement d’exécution
description: Charger le DOM et l’environnement d’exécution des compléments Office
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 557297fc9e13ebab5b4eebd7917d0e0d9e444e88
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608124"
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

4. Le contrôle de navigateur charge le DOM et le corps HTML, et appelle le gestionnaire d’événements pour l' `window.onload` événement.

5. L’application hôte Office charge l’environnement d’exécution, lequel télécharge et met en cache l’API JavaScript pour les fichiers de bibliothèque JavaScript à partir du serveur de réseau de distribution de contenu, puis appelle le gestionnaire d’événements du complément pour l’événement [initialize](/javascript/api/office#office-initialize-reason-) de l’objet [Office](/javascript/api/office) si un gestionnaire lui a été affecté. Il vérifie alors également si des rappels (ou des fonctions `then()` chaînées) ont été transmis (ou chaînées) au gestionnaire `Office.onReady`. Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady` , voir [initialiser votre complément](initialize-add-in.md).

6. Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.


## <a name="startup-of-an-outlook-add-in"></a>Démarrage d’un complément Outlook

La figure suivante illustre le flux des événements impliqués au démarrage d’un complément Outlook exécuté sur un ordinateur de bureau, une tablette ou un smartphone.

![Flux des événements au démarrage du complément Outlook](../images/outlook15-loading-dom-agave-runtime.png)

Les événements suivants se produisent lors du démarrage d’un complément Outlook :

1. Lorsqu’Outlook démarre, il lit les manifestes XML pour les compléments Outlook qui ont été installés pour le compte de messagerie de l’utilisateur.

2. L’utilisateur sélectionne un élément dans Outlook.

3. Si l’élément sélectionné répond aux conditions d’activation d’un complément Outlook, Outlook active le complément et affiche son bouton dans l’interface utilisateur.

4. Si l’utilisateur clique sur le bouton pour démarrer le complément Outlook, Outlook ouvre la page HTML dans un contrôle de navigateur. Les deux étapes suivantes, 5 et 6, se produisent en parallèle.

5. Le contrôle de navigateur charge le DOM et le corps HTML, et appelle le gestionnaire d’événements pour l' `onload` événement.

6. Outlook charge l’environnement d’exécution, lequel télécharge et met en cache l’API JavaScript pour les fichiers de bibliothèque JavaScript à partir du serveur de réseau de distribution de contenu, puis appelle le gestionnaire d’événements du complément pour l’événement [initialize](/javascript/api/office#office-initialize-reason-) de l’objet [Office](/javascript/api/office) du complément si un gestionnaire lui a été affecté. Il vérifie alors également si des rappels (ou des fonctions `then()` chaînées) ont été transmis (ou chaînées) au gestionnaire `Office.onReady`. Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady` , voir [initialiser votre complément](initialize-add-in.md).

7. Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.


## <a name="checking-the-load-status"></a>Vérification du statut de chargement

Vous pouvez vérifier que le chargement du DOM et de l’environnement d’exécution est bien terminé en utilisant la fonction jQuery [.ready()](https://api.jquery.com/ready/) : `$(document).ready()`. Par exemple, le `onReady` Gestionnaire d’événements suivant vérifie que le DOM est chargé pour la première fois avant l’exécution du code spécifique à l’initialisation du complément. Par la suite, le `onReady` Gestionnaire continue d’utiliser la propriété [Mailbox. Item](/javascript/api/outlook/office.mailbox#item) pour obtenir l’élément actuellement sélectionné dans Outlook et appelle la fonction principale du complément, `initDialer` .

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

Vous pouvez également utiliser le même code dans un gestionnaire d' `initialize` événements comme illustré dans l’exemple suivant.

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

Cette même technique peut être utilisée dans les `onReady` `initialize` gestionnaires ou des compléments Office.

Le numéroteur téléphonique fourni comme exemple de complément Outlook présente une approche légèrement différente, puisqu’il utilise uniquement JavaScript pour vérifier ces mêmes conditions.

> [!IMPORTANT]
> Même si aucune tâche d’initialisation n’est à effectuer dans votre complément, vous devez inclure au moins un appel de `Office.onReady` `Office.initialize` la fonction de gestionnaire d’événements minimal, comme illustré dans les exemples suivants.
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> Si vous n’appelez pas `Office.onReady` ou n’assignez pas de `Office.initialize` Gestionnaire d’événements, votre complément peut déclencher une erreur lors de son démarrage. En outre, si un utilisateur essaie d’utiliser votre complément avec un client web Office, notamment Excel, PowerPoint ou Outlook, l’exécution du complément échouera.
>
> Si votre complément comprend plusieurs pages, chaque fois qu’il charge une nouvelle page, celle-ci doit appeler `Office.onReady` ou assigner un `Office.initialize` Gestionnaire d’événements.

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [Initialiser votre complément Office](initialize-add-in.md)
