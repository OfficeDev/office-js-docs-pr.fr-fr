---
title: Chargement du DOM et de l’environnement d’exécution
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 3ce0da16a134c435147f7106d6bea9c006ce2922
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944047"
---
# <a name="loading-the-dom-and-runtime-environment"></a>Chargement du DOM et de l’environnement d’exécution



Un complément doit s’assurer que le DOM et l’environnement d’exécution des Compléments Office ont été chargés avant d’exécuter sa propre logique personnalisée. 

## <a name="startup-of-a-content-or-task-pane-add-in"></a>Démarrage d’un complément de contenu ou du volet Office

La figure suivante illustre le flux des événements impliqués au démarrage d’un complément de contenu ou du volet Office dans Excel, PowerPoint, Project, Word ou Access.

![Flux des événements au démarrage d’un complément de contenu ou du volet Office](../images/office15-app-sdk-loading-dom-agave-runtime.png)

Les événements suivants se produisent lors du démarrage d’un complément de contenu ou du volet Office : 



1. L’utilisateur ouvre un document qui contient déjà un complément ou insère un complément dans le document.
    
2. L’application hôte Office lit le manifeste XML du complément à partir d’AppSource, d’un catalogue de compléments sur SharePoint ou du catalogue de dossiers partagés duquel il provient.
    
3. L’application hôte Office ouvre la page HTML du complément dans un contrôle de navigateur.
    
    Les deux étapes suivantes, 4 et 5, se produisent de manière asynchrone et parallèlement. C’est pour cela que le code de votre complément doit veiller à ce que le chargement du DOM et de l’environnement d’exécution du complément soit terminé avant de continuer.
    
4. Le contrôle de navigateur charge le DOM et le corps HTML, puis demande au gestionnaire d’événements l’événement  **window.onload**.
    
5. L’application hôte Office charge l’environnement d’exécution, lequel télécharge et met en cache l’API JavaScript pour les fichiers de bibliothèque JavaScript à partir du serveur de réseau de distribution de contenu, puis appelle le gestionnaire d’événements du complément pour l’événement [initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) de l’objet [Office](https://docs.microsoft.com/javascript/api/office?view=office-js).
    
6. Lorsque le chargement du modèle objet de document (DOM) et du corps HTML est terminé et que le complément s’est initialisé, la fonction principale de l’application peut s’exécuter.
    

## <a name="startup-of-an-outlook-add-in"></a>Démarrage d’un complément Outlook



La figure suivante illustre le flux des événements impliqués au démarrage d’un complément Outlook exécuté sur un ordinateur de bureau, une tablette ou un smartphone.

![Flux des événements au démarrage du complément Outlook](../images/outlook15-loading-dom-agave-runtime.png)

Les événements suivants se produisent lors du démarrage d’un complément Outlook : 



1. Lorsqu’Outlook démarre, il lit les manifestes XML pour les compléments Outlook qui ont été installés pour le compte de messagerie de l’utilisateur.
    
2. L’utilisateur sélectionne un élément dans Outlook.
    
3. Si l’élément sélectionné répond aux conditions d’activation d’un complément Outlook, Outlook active le complément et affiche son bouton dans l’interface utilisateur.
    
4. Si l’utilisateur clique sur le bouton pour démarrer le complément Outlook, Outlook ouvre la page HTML dans un contrôle de navigateur. Les deux étapes suivantes, 5 et 6, se produisent en parallèle.
    
5. Le contrôle de navigateur charge le modèle objet de document (DOM) et le corps HTML, puis appelle le gestionnaire d’événements pour l’événement  **onload**.
    
6. Outlook appelle le gestionnaire d’événements pour l’événement [initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) de l’objet [Office](https://docs.microsoft.com/javascript/api/office?view=office-js) du complément.
    
7. Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.
    

## <a name="checking-the-load-status"></a>Vérification du statut de chargement


Pour vérifier que le chargement du modèle objet de document (DOM) et de l’environnement d’exécution des est terminé, il est notamment possible d’utiliser la fonction jQuery [.ready()](http://api.jquery.com/ready/) :  `$(document).ready()`. Par exemple, la fonction de gestionnaire d’événements  **initialize** ci-dessous s’assure d’abord que le DOM est bien chargé avant l’exécution du code d’initialisation du complément. Par conséquent, le gestionnaire d’événements **initialize** utilise la propriété [mailbox.item](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) pour obtenir l’élément actuellement sélectionné dans Outlook, puis appelle la fonction principale du complément, `initDialer`.


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

Il est possible d’utiliser cette même technique dans le gestionnaire  **initialize** de toute Complément Office.

Le numéroteur téléphonique fourni comme exemple de complément Outlook présente une approche légèrement différente, puisqu’il utilise uniquement JavaScript pour vérifier ces mêmes conditions. 

> [!IMPORTANT]
> Même si aucune tâche d’initialisation n’est à effectuer dans votre complément, vous devez inclure au moins une fonction de gestionnaire d’événements **Office.initialize** minimale comme l’exemple suivant.

```js
Office.initialize = function () {
};
```

Si vous n’incluez pas de gestionnaire d’événements  **Office.initialize**, votre complément peut générer une erreur au démarrage. En outre, si un utilisateur tente d’utiliser votre complément avec un client web Office Online, comme Excel Online, PowerPoint Online ou Outlook Web App, il n’est pas exécuté.

Si votre complément comprend plusieurs pages, chaque fois qu’il charge une nouvelle page, celle-ci doit inclure ou appeler un gestionnaire d’événements  **Office.initialize**.


## <a name="see-also"></a>Voir aussi

- [Présentation de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
    
