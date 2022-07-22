---
title: Chargement du DOM et de l’environnement d’exécution
description: Chargez l’environnement d’exécution dom et des compléments Office.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: be93b261c8beacdb7b4e8cd08448abf06b14607e
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958684"
---
# <a name="loading-the-dom-and-runtime-environment"></a>Chargement du DOM et de l’environnement d’exécution

Un complément doit s’assurer que le DOM et l’environnement d’exécution des Compléments Office ont été chargés avant d’exécuter sa propre logique personnalisée.

## <a name="startup-of-a-content-or-task-pane-add-in"></a>Démarrage d’un complément de contenu ou du volet Office

La figure suivante illustre le flux des événements impliqués au démarrage d’un complément de contenu ou du volet Office dans Excel, PowerPoint, Project ou Word.

![Flux d’événements lors du démarrage d’un complément de contenu ou de volet Office.](../images/office15-app-sdk-loading-dom-agave-runtime.png)

Les événements suivants se produisent lorsqu’un complément de contenu ou de volet Office démarre.

1. L’utilisateur ouvre un document qui contient déjà un complément ou insère un complément dans le document.

2. L’application cliente Office lit le manifeste XML du complément à partir d’AppSource, d’un catalogue d’applications sur SharePoint ou du catalogue de dossiers partagés dont il provient.

3. L’application cliente Office ouvre la page HTML du complément dans un contrôle de navigateur.

    Les deux étapes suivantes, 4 et 5, se produisent de manière asynchrone et parallèlement. C’est pour cela que le code de votre complément doit veiller à ce que le chargement du DOM et de l’environnement d’exécution du complément soit terminé avant de continuer.

4. Le contrôle de navigateur charge le corps DOM et HTML, puis appelle le gestionnaire d’événements pour l’événement `window.onload` .

5. L’application cliente Office charge l’environnement d’exécution, qui télécharge et met en cache les fichiers de bibliothèque d’API JavaScript Office à partir du serveur du réseau de distribution de contenu (CDN), puis appelle le gestionnaire d’événements du complément pour l’événement [d’initialisation](/javascript/api/office#Office_initialize_reason_) de l’objet [Office](/javascript/api/office) , si un gestionnaire lui a été attribué. À ce stade, il vérifie également si des rappels (ou une méthode chaînée `then()` ) ont été passés (ou chaînés) au `Office.onReady` gestionnaire. Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady`, consultez [Initialiser votre complément](initialize-add-in.md).

6. Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.

## <a name="startup-of-an-outlook-add-in"></a>Démarrage d’un complément Outlook

La figure suivante illustre le flux des événements impliqués au démarrage d’un complément Outlook exécuté sur un ordinateur de bureau, une tablette ou un smartphone.

![Flux d’événements lors du démarrage du complément Outlook.](../images/outlook15-loading-dom-agave-runtime.png)

Les événements suivants se produisent lorsqu’un complément Outlook démarre.

1. Lorsqu’Outlook démarre, il lit les manifestes XML pour les compléments Outlook qui ont été installés pour le compte de messagerie de l’utilisateur.

2. L’utilisateur sélectionne un élément dans Outlook.

3. Si l’élément sélectionné répond aux conditions d’activation d’un complément Outlook, Outlook active le complément et affiche son bouton dans l’interface utilisateur.

4. Si l’utilisateur clique sur le bouton pour démarrer le complément Outlook, Outlook ouvre la page HTML dans un contrôle de navigateur. Les deux étapes suivantes, 5 et 6, se produisent en parallèle.

5. Le contrôle de navigateur charge le corps DOM et HTML, puis appelle le gestionnaire d’événements pour l’événement `onload` .

6. Outlook charge l’environnement d’exécution, lequel télécharge et met en cache l’API JavaScript pour les fichiers de bibliothèque JavaScript à partir du serveur de réseau de distribution de contenu, puis appelle le gestionnaire d’événements du complément pour l’événement [initialize](/javascript/api/office#Office_initialize_reason_) de l’objet [Office](/javascript/api/office) du complément si un gestionnaire lui a été affecté. À ce stade, il vérifie également si des rappels (ou méthodes chaînées `then()` ) ont été passés (ou chaînés) au `Office.onReady` gestionnaire. Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady`, consultez [Initialiser votre complément](initialize-add-in.md).

7. Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [Initialiser votre complément Office](initialize-add-in.md)
