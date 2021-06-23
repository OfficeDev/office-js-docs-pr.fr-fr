---
title: Chargement du DOM et de l’environnement d’exécution
description: Chargez le DOM et Office’environnement d’runtime des add-ins.
ms.date: 04/20/2021
localization_priority: Normal
ms.openlocfilehash: e66e6d5e30f5305dce35157280210a371ee3896f
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076012"
---
# <a name="loading-the-dom-and-runtime-environment"></a>Chargement du DOM et de l’environnement d’exécution

Un complément doit s’assurer que le DOM et l’environnement d’exécution des Compléments Office ont été chargés avant d’exécuter sa propre logique personnalisée.

## <a name="startup-of-a-content-or-task-pane-add-in"></a>Démarrage d’un complément de contenu ou du volet Office

La figure suivante illustre le flux des événements impliqués au démarrage d’un complément de contenu ou du volet Office dans Excel, PowerPoint, Project ou Word.

![Flow événements lors du démarrage d’un module de contenu ou du volet Des tâches.](../images/office15-app-sdk-loading-dom-agave-runtime.png)

Les événements suivants se produisent lors du démarrage d’un complément de contenu ou du volet Office :

1. L’utilisateur ouvre un document qui contient déjà un complément ou insère un complément dans le document.

2. L’application cliente Office lit le manifeste XML du add-in à partir d’AppSource, d’un catalogue d’applications sur SharePoint ou du catalogue de dossiers partagés dont il est issu.

3. L Office application cliente ouvre la page HTML du module dans un contrôle de navigateur.

    Les deux étapes suivantes, 4 et 5, se produisent de manière asynchrone et parallèlement. C’est pour cela que le code de votre complément doit veiller à ce que le chargement du DOM et de l’environnement d’exécution du complément soit terminé avant de continuer.

4. Le contrôle de navigateur charge le DOM et le corps HTML, puis appelle le responsable de l’événement `window.onload` pour l’événement.

5. L’application cliente Office charge l’environnement d’utilisation, qui télécharge et met en cache les fichiers de bibliothèque d’API JavaScript Office à partir du serveur de réseau de distribution de contenu (CDN), puis appelle le responsable des événements du module pour [l’événement d’initialisation](/javascript/api/office#office-initialize-reason-) de l’objet [Office,](/javascript/api/office) si un handler lui a été affecté. Il vérifie alors également si des rappels (ou des fonctions `then()` chaînées) ont été transmis (ou chaînées) au gestionnaire `Office.onReady`. Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady` , voir [Initialiser votre add-in](initialize-add-in.md).

6. Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.


## <a name="startup-of-an-outlook-add-in"></a>Démarrage d’un complément Outlook

La figure suivante illustre le flux des événements impliqués au démarrage d’un complément Outlook exécuté sur un ordinateur de bureau, une tablette ou un smartphone.

![Flow d’événements au démarrage Outlook de votre module.](../images/outlook15-loading-dom-agave-runtime.png)

Les événements suivants se produisent lors du démarrage d’un complément Outlook :

1. Lorsqu’Outlook démarre, il lit les manifestes XML pour les compléments Outlook qui ont été installés pour le compte de messagerie de l’utilisateur.

2. L’utilisateur sélectionne un élément dans Outlook.

3. Si l’élément sélectionné répond aux conditions d’activation d’un complément Outlook, Outlook active le complément et affiche son bouton dans l’interface utilisateur.

4. Si l’utilisateur clique sur le bouton pour démarrer le complément Outlook, Outlook ouvre la page HTML dans un contrôle de navigateur. Les deux étapes suivantes, 5 et 6, se produisent en parallèle.

5. Le contrôle de navigateur charge le DOM et le corps HTML, puis appelle le responsable de l’événement `onload` pour l’événement.

6. Outlook charge l’environnement d’exécution, lequel télécharge et met en cache l’API JavaScript pour les fichiers de bibliothèque JavaScript à partir du serveur de réseau de distribution de contenu, puis appelle le gestionnaire d’événements du complément pour l’événement [initialize](/javascript/api/office#office-initialize-reason-) de l’objet [Office](/javascript/api/office) du complément si un gestionnaire lui a été affecté. Il vérifie alors également si des rappels (ou des fonctions `then()` chaînées) ont été transmis (ou chaînées) au gestionnaire `Office.onReady`. Pour plus d’informations sur la distinction entre `Office.initialize` et `Office.onReady` , voir [Initialiser votre add-in](initialize-add-in.md).

7. Lorsque le chargement du DOM et du corps HTML est terminé et que le complément finit de s’initialiser, la fonction principale du complément peut poursuivre.

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [Initialiser votre complément Office](initialize-add-in.md)
