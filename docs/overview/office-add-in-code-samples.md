---
title: Exemples de code Office
description: Une liste d Office exemples de code pour vous aider à apprendre et à créer vos propres modules.
ms.date: 11/18/2021
localization_priority: high
ms.openlocfilehash: 62bce4c5c6a01895db5d183e0526e0af820156ea
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/15/2021
ms.locfileid: "61514012"
---
# <a name="office-add-in-code-samples"></a>Exemples de code Office

Ces exemples de code sont écrits pour vous aider à apprendre à utiliser différentes fonctionnalités lors du développement de Office de développement.

## <a name="getting-started"></a>Prise en main

Les exemples suivants montrent comment créer le complément Office le plus simple avec uniquement un manifeste, une page web HTML et un logo. Ces composants sont les éléments fondamentaux d’un complément Office. Pour plus d’informations sur la prise en main, consultez nos [démarrages rapides](../quickstarts/excel-quickstart-jquery.md) et [didacticiels](/search/?terms=tutorial&scope=Office%20Add-ins).

* [Complément Excel « Hello World »](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/hello-world/excel-hello-world)
* [Complément Outlook « Hello World »](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/hello-world/outlook-hello-world)
* [Complément PowerPoint « Hello World »](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/hello-world/powerpoint-hello-world)
* [Complément Word « Hello World »](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/hello-world/word-hello-world)

## <a name="outlook"></a>Outlook

| Nom                | Description         |
|:--------------------|:--------------------|
| [Utiliser l’activation Outlook basée sur un événement pour marquer des destinataires externes (aperçu)](/samples/officedev/pnp-officeaddins/outlook-add-in-tag-external-recipients) | Utilisez l’activation basée sur des événements pour exécuter Outlook complément lorsque l’utilisateur modifie des destinataires lors de la composition d’un message. Le complément utilise également `appendOnSendAsync` l’API pour ajouter une clause d’exclusion de responsabilité. |
| [Utiliser l’activation Outlook basée sur un événement pour définir la signature](/samples/officedev/pnp-officeaddins/outlook-add-in-set-signature/) | Utilisez l'activation basée sur des événements pour exécuter un module complémentaire Outlook lorsque l'utilisateur crée un nouveau message ou un rendez-vous. Le module peut répondre aux événements, même lorsque le volet Tâches n’est pas ouvert. Il utilise également `setSignatureAsync` l’API. |

## <a name="excel"></a>Excel

| Nom                | Description         |
|:--------------------|:--------------------|
| [Ouvrir dans Teams](/samples/officedev/pnp-officeaddins/office-excel-add-in-open-in-teams/) | Créez une feuille Excel feuille de calcul Microsoft Teams contenant les données que vous définissez.|
| [Insérez un fichier Excel externe et remplissez-le avec des données JSON](/samples/officedev/pnp-officeaddins/excel-add-in-insert-external-file/)  | Insérez un modèle existant à partir d'un fichier Excel externe dans le classeur Excel actuellement ouvert. Ensuite, remplissez le modèle avec les données d'un service Web JSON. |
| [Créer des onglets contextuels personnalisés sur le ruban](/samples/officedev/pnp-officeaddins/office-add-in-contextual-tabs/) | Créez un onglet contextuel personnalisé sur le ruban dans l’interface de l’utilisateur Office. L’exemple crée un tableau et lorsque l’utilisateur déplace le focus à l’intérieur du tableau, l’onglet personnalisé s’affiche. Lorsque l’utilisateur se déplace en dehors du tableau, l’onglet personnalisé est masqué. |
| [Utiliser des raccourcis clavier pour les actions Office de la recherche](/samples/officedev/pnp-officeaddins/office-add-in-keyboard-shortcuts) | Configurer un projet de Excel de base qui utilise des raccourcis clavier |
| [Exemple de fonction personnalisée utilisant le service web](/samples/officedev/pnp-officeaddins/excel-custom-function-web-worker-pattern/) | Utilisez les web workers dans les fonctions personnalisées pour éviter de bloquer l'interface utilisateur de votre module complémentaire Office. |
| [Utiliser des techniques de stockage pour accéder aux données à partir d’un Office lorsqu’il est hors connexion](/samples/officedev/pnp-officeaddins/use-storage-techniques-to-access-data-from-an-office-add-in-when-offline/) | Implémentez le stockage local pour activer des fonctionnalités limitées pour votre Office lorsqu’un utilisateur subit une perte de connexion. |
| [Modèle de traitement par lots de fonctions personnalisées](/samples/officedev/pnp-officeaddins/excel-custom-function-batching-pattern/)| Traitement par lots de plusieurs appels en un seul appel pour réduire le nombre d’appels réseau vers un service distant.|

## <a name="shared-javascript-runtime"></a>Runtime JavaScript partagé

| Nom                | Description         |
|:--------------------|:--------------------|
[Partager des données globales avec un runtime partagé](/samples/officedev/pnp-officeaddins/office-add-in-shared-runtime-global-data/) | Configurer un projet de base qui utilise le runtime partagé pour exécuter le code pour les boutons du ruban, le volet Des tâches et les fonctions personnalisées dans un seul runtime de navigateur. |
| [Gérer l’interface utilisateur du ruban et du volet Des tâches, et exécuter le code sur le document ouvert](/samples/officedev/pnp-officeaddins/office-add-in-ribbon-task-pane-ui/) | Créez des boutons de ruban contextuels qui sont activés en fonction de l’état de votre complément. |

## <a name="authentication-authorization-and-single-sign-on-sso"></a>Authentification, autorisation et authentification unique (SSO)

| Nom                | Description         |
|:--------------------|:--------------------|
| [Exemple d' sign-on (SSO) Outlook de l' sign-on unique (SSO)](/samples/officedev/pnp-officeaddins/outlook-add-in-sso-aspnet/) | Utilisez la fonction SSO d'Office pour permettre à l'extension d'accéder aux données Microsoft Graph.|
| [Obtenir des données OneDrive à l’aide de Microsoft Graph et msal.js dans un complément Office](/samples/officedev/pnp-officeaddins/office-add-in-auth-graph-react/) | Créez un module complémentaire Office, en tant qu'application monopage (SPA) sans backend, qui se connecte à Microsoft Graph et accède aux classeurs stockés dans OneDrive Entreprise pour mettre à jour une feuille de calcul.  |
| [Authentification du complément Office à Microsoft Graph](/samples/officedev/pnp-officeaddins/office-add-in-auth-aspnet-graph/) | Apprenez à créer un complément Microsoft Office qui se connecte à Microsoft Graph, et à accéder aux classeurs stockés dans OneDrive Entreprise pour mettre à jour une feuille de calcul. |
| [Autorisation du module d'extension Outlook pour Microsoft Graph](/samples/officedev/pnp-officeaddins/outlook-add-in-auth-aspnet-graph/). | Créez un module complémentaire Outlook qui se connecte à Microsoft Graph et accède aux classeurs stockés dans OneDrive Entreprise pour composer un nouveau message électronique. |
| [Sign-on (SSO) Office add-in with ASP.NET](/samples/officedev/pnp-officeaddins/office-add-in-sso-aspnet/) | Utilisez `getAccessToken` l'API dans Office.js pour donner au complément un accès aux données Microsoft Graph. Cet exemple est construit sur ASP.NET. |
| [Sign-on (SSO) Office add-in with Node.js](/samples/officedev/pnp-officeaddins/office-add-in-sso-nodejs/) | Utilisez `getAccessToken` l'API dans Office.js pour donner au complément un accès aux données Microsoft Graph. Cet exemple est construit sur Node.js.|

## <a name="additional-samples"></a>Exemples supplémentaires

| Nom                | Description         |
|:--------------------|:--------------------|
|[Utiliser une bibliothèque partagée pour migrer votre Visual Studio Tools pour Office vers un Office web](/samples/officedev/pnp-officeaddins/vsto-shared-library-excel/) |Fournit une stratégie pour la réutilisation du code lors de la migration de VSTO vers Office de code. |
| [Intégrer une fonction Azure à votre Excel personnalisée](/samples/officedev/pnp-officeaddins/azure-function-with-excel-custom-function/) | Intégrez des fonctions Azure à des fonctions personnalisées pour passer au cloud ou intégrer des services supplémentaires. |
|[Exemples de code DPI dynamique](/samples/officedev/pnp-officeaddins/dynamic-dpi-code-samples/) |Une collection d’exemples pour la gestion des modifications de DPI dans COM, VSTO et Office des compléments. |

## <a name="next-steps"></a>Étapes suivantes

Rejoignez le programme pour développeurs Microsoft 365. Obtenez un bac à sable gratuit, des outils et d'autres ressources dont vous avez besoin pour créer des solutions pour la plate-forme Microsoft 365.

- [Bac à sable développeur gratuit](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Obtenez un abonnement gratuit et renouvelable de 90 jours Microsoft 365 E5 développeur.
- [Packs d’exemples de données](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Configurez automatiquement votre bac à sable en installant les données utilisateur et le contenu pour vous aider à créer vos solutions.
- [Accès aux experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Accéder aux événements de la communauté pour en savoir plus Microsoft 365 experts.
- [Recommandations personnalisées ](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations)Trouvez rapidement des ressources pour les développeurs depuis votre tableau de bord personnalisé.
