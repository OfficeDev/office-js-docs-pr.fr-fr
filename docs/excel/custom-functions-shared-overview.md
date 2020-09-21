---
ms.date: 08/13/2020
description: Découvrez l'exécution de fonctions personnalisées, les boutons du ruban et le code du volet des tâches dans un runtime JavaScript identique pour coordonner des scénarios dans votre complément.
title: Exécutez votre code de complément dans un runtime JavaScript partagé
localization_priority: Priority
ms.openlocfilehash: 04932bcf292686fd9d0abf2ff99c19f062f21456
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819545"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtimes"></a>Vue d’ensemble : exécutez votre code de complément dans un runtime JavaScript partagé

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Lors de l’exécution d’Excel sur Windows ou Mac, votre complément exécute le code des boutons du ruban, des fonctions personnalisées et du volet des tâches dans des environnements runtime JavaScript distincts. Cela permet de créer des limitations, telles que l'impossibilité de partager aisément des données globales ou de pouvoir accéder à l'ensemble des fonctionnalités CORS à partir d’une fonction personnalisée.

Vous pouvez toutefois configurer votre complément Excel pour partager un code dans le même runtime JavaScript (également appelé runtime partagé). Vous pouvez ainsi améliorer la coordination dans votre complément et accéder au volet des tâches DOM et CORS à partir de toutes les parties de votre complément.

La configuration d’un runtime partagé permet les scénarios suivants :

- Votre complément dispose d'un DOM partagé auquel le ruban, le volet des tâches et les fonctions personnalisées peuvent accéder.
- Vos fonctions personnalisées bénéficieront d'une prise en charge complète de CORS.
- Vos fonctions personnalisées peuvent appeler les API Office.js pour lire les données d’un document feuille de calcul.
- Votre complément peut exécuter un code dès que le document est ouvert.
- Votre complément peut continuer à exécuter un code lorsque le volet des tâches est fermé.

Lorsque vous exécutez des fonctions personnalisées dans une exécution partagée à l’aide du volet Office, celui-ci s’exécute dans une instance de navigateur Microsoft Internet Explorer 11, comme expliqué dans [navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md). De plus, les boutons affichés par votre complément Excel sur le ruban s’exécutent dans le même runtime partagé. L’image ci-après présente l'exécution des fonctions personnalisées, de interface utilisateur du ruban et du code du volet des tâches dans le même runtime JavaScript.

![Fonctions personnalisées s'exécutant dans un runtime partagé avec les boutons du ruban et le volet des tâches dans Excel](../images/custom-functions-in-browser-runtime.png)

## <a name="set-up-a-shared-runtime"></a>Configurer un runtime partagé

Pour plus d’informations sur la configuration de vos fonctions personnalisées de manière à utiliser un runtime partagé, voir la [configuration d’un article d’exécution partagé](configure-your-add-in-to-use-a-shared-runtime.md) .

### <a name="debugging"></a>Débogage

Lors de l’utilisation d’un runtime partagé, vous ne pouvez pas utiliser Visual Studio Code pour déboguer des fonctions personnalisées dans Excel sur Windows à cette date. Vous devez utiliser les outils de développement à la place. Pour plus d'informations, voir le [Débogage des compléments avec les outils de développement sur Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).

## <a name="give-us-feedback"></a>Faites-nous part de vos commentaires

Nous aimerions connaître votre avis concernant cette fonctionnalité. Si vous trouvez des bogues, des problèmes ou si vous avez des questions relatives à cette fonctionnalité, faites-le nous savoir en créant un problème GitHub dans le [référentiel Office-js](https://github.com/OfficeDev/office-js).

## <a name="see-also"></a>Voir aussi

- [Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Appeler des API Excel à partir de votre fonction personnalisée](call-excel-apis-from-custom-function.md)
