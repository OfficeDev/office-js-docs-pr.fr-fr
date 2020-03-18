---
ms.date: 02/13/2020
description: Découvrez l'exécution de fonctions personnalisées, les boutons du ruban et le code du volet des tâches dans un runtime JavaScript identique pour coordonner des scénarios dans votre complément.
title: Exécutez votre code de complément dans un runtime JavaScript partagé (préversion)
localization_priority: Priority
ms.openlocfilehash: 774990a9452d450bd5c4d968027bc64ebee858af
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719531"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtime-preview"></a>Vue d’ensemble : exécutez votre code de complément dans un runtime JavaScript partagé (préversion)

[!include[Running custom functions in shared JavaScript runtime note](../includes/excel-shared-runtime-preview-note.md)]

Lors de l’exécution d’Excel sur Windows ou Mac, votre complément exécute le code des boutons du ruban, des fonctions personnalisées et du volet des tâches dans des environnements runtime JavaScript distincts. Cela permet de créer des limitations, telles que l'impossibilité de partager aisément des données globales ou de pouvoir accéder à l'ensemble des fonctionnalités CORS à partir d’une fonction personnalisée.

Vous pouvez toutefois configurer votre complément Excel pour partager un code dans le même runtime JavaScript (également appelé runtime partagé). Vous pouvez ainsi améliorer la coordination dans votre complément et accéder au volet des tâches DOM et CORS à partir de toutes les parties de votre complément.

La configuration d’un runtime partagé permet les scénarios suivants :

- Votre complément dispose d'un DOM partagé auquel le ruban, le volet des tâches et les fonctions personnalisées peuvent accéder.
- Vos fonctions personnalisées bénéficieront d'une prise en charge complète de CORS.
- Vos fonctions personnalisées peuvent appeler les API Office.js pour lire les données d’un document feuille de calcul.
- Votre complément peut exécuter un code dès que le document est ouvert.
- Votre complément peut continuer à exécuter un code lorsque le volet des tâches est fermé.

Lorsque vous exécutez des fonctions personnalisées dans un runtime partagé avec le volet des tâches, celui-ci s’exécute dans une instance de navigateur sur différentes plateformes, tel qu'expliqué dans [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md). En outre, les boutons affichés sur le ruban par votre complément Excel s’exécutent dans le même runtime partagé. L’image ci-après présente l'exécution des fonctions personnalisées, de interface utilisateur du ruban et du code du volet des tâches dans le même runtime JavaScript.

![Fonctions personnalisées s'exécutant dans le runtime partagé avec les boutons du ruban et le volet des tâches dans Excel](../images/custom-functions-in-browser-runtime.png)

## <a name="differences-when-running-custom-functions-in-a-shared-runtime"></a>Différences lors de l’exécution de fonctions personnalisées dans un runtime partagé

Lorsque vous configurez votre projet de complément Excel pour l’exécution de fonctions personnalisées dans un runtime partagé, il existe quelques différences dans l’utilisation de la fonction runtime personnalisée.

### <a name="storage"></a>Stockage

Vous n’avez plus besoin d’utiliser l’API de **Stockage** pour partager des données entre le volet des tâches, les fonctions personnalisées ou l’interface utilisateur du ruban. Vous pouvez placer des variables globales dans l'objet de la **fenêtre**, ou utiliser votre propre approche de gestion d'état préférée.

### <a name="authentication"></a>Authentification

Lorsque vous recevez des jetons dans le cadre de l’authentification, vous n’avez pas besoin d’utiliser l’API de **stockage** pour les partager avec le volet des tâches, les fonctions personnalisées et l’interface utilisateur du ruban. Vous pouvez utiliser votre propre technique de stockage préférée par défaut et un emplacement de stockage pour les partager, tel que `localStorage`.

### <a name="dialog-api"></a>API de boîte de dialogue

Vous n’avez plus besoin d’utiliser l'API **OfficeRuntime.Dialog** pour afficher une boîte de dialogue à partir d’une fonction personnalisée. Vous pouvez utiliser la même API [boîte de dialogue](../develop/dialog-api-in-office-add-ins.md) pour les fonctions personnalisées, les boutons du ruban et le volet des tâches.

### <a name="debugging"></a>Débogage

Lors de l’utilisation d’un runtime partagé, vous ne pouvez pas utiliser Visual Studio Code pour déboguer des fonctions personnalisées dans Excel sur Windows à cette date. Des outils de développeur sont nécessaires. Pour plus d'informations, voir le [Débogage des compléments avec les outils de développement sur Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).

## <a name="get-started"></a>Prise en main

Pour configurer votre projet de complément Excel pour l’exécution de fonctions personnalisées dans un runtime partagé, voir [Configurer votre complément Excel pour utiliser un runtime JavaScript partagé (préversion)](configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="give-us-feedback"></a>Faites nous part de vos commentaires

Nous aimerions connaître votre avis concernant cette fonctionnalité. Si vous trouvez des bogues, des problèmes ou si vous avez des questions relatives à cette fonctionnalité, faites-le nous savoir en créant un problème GitHub dans le [référentiel Office-js](https://github.com/OfficeDev/office-js).

## <a name="see-also"></a>Voir aussi

Listes des articles associés relatif au runtime partagé
- [Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office (préversion)](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Appeler des API Excel à partir de votre fonction personnalisée (préversion)](call-excel-apis-from-custom-function.md)