---
title: Débogage de compléments dans Office sur le web
description: Découvrez comment utiliser Office sur le web pour tester et déboguer vos compléments.
ms.date: 03/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5a07185c064d65432c7a3afce1e9f32e99034c3e
ms.sourcegitcommit: 3d7792b1f042db589edb74a895fcf6d7ced63903
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2022
ms.locfileid: "63435689"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Débogage de compléments dans Office sur le web

Cet article explique comment utiliser Office sur le Web pour déboguer vos modules. Utilisez cette technique :

- Pour déboguer des applications sur un ordinateur qui n’exécute pas Windows ou le client&mdash; de bureau Office par exemple, si vous développez sur un Mac ou Linux.
- Autre processus de débogage si vous ne pouvez pas ou ne le souhaitez pas, déboguer dans un IDE, tel que Visual Studio ou Visual Studio Code.

Cet article suppose que vous avez un projet de add-in qui doit être déboité. Si vous souhaitez simplement pratique le débogage sur le web, créez un projet à l’aide de l’un des démarrages rapides pour des applications Office spécifiques, telles que ce démarrage rapide [pour Word](../quickstarts/word-quickstart.md).

## <a name="debug-your-add-in"></a>Déboguer votre complément

Pour déboguer votre complément à l’aide d’Office sur le web, procédez comme suit :

1. Exécutez le projet sur localhost et chargez-le dans un document dans Office sur le Web. Pour obtenir des instructions détaillées sur le chargement d’une version de version Office des applications sur [le web](sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web-manually).

2. Ouvrez les outils de développement du navigateur. Pour ce faire, il s’agit généralement d’appuyer sur F12. Ouvrez l’outil débogger et utilisez-le pour définir des points d’arrêt et observer des variables. Pour obtenir de l’aide détaillée sur l’utilisation de l’outil de votre navigateur, consultez l’une des informations suivantes.  

   - [Firefox](https://developer.mozilla.org/en-US/docs/Tools)
   - [Safari](https://support.apple.com/guide/safari/use-the-developer-tools-in-the-develop-menu-sfri20948/mac)
   - [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
   - [Déboguer des compléments à l’aide des outils de développement pour la version héritée Edge](debug-add-ins-using-devtools-edge-legacy.md)

   > [!NOTE]
   > Office sur le Web ne s’ouvre pas dans Internet Explorer.

## <a name="potential-issues"></a>Problèmes potentiels

Voici quelques problèmes que vous pouvez rencontrer lors du débogage.

- Certaines erreurs JavaScript peuvent provenir d’Office sur le web.

- Le navigateur peut afficher une erreur relative à un certificat non valide que vous devrez contourner. Le processus d’exécution de cette opération varie en fonction du navigateur et des interfaces utilisateur des différents navigateurs permettant d’effectuer cette modification régulièrement. Vous devez effectuer une recherche dans l’aide du navigateur ou rechercher des instructions en ligne. (Par exemple, recherchez « Avertissement de certificat Microsoft Edge non valide ».) La plupart des navigateurs, sur la page d’avertissement, comportent un lien qui vous permet d’accéder à la page du complément. Par exemple, Microsoft Edge comporte un lien « Accéder à la page web (non recommandé) ». En général, vous devez passer par ce lien chaque fois que le complément est rechargé. Pour un contournement plus long, consultez l’aide comme suggéré.

- Si vous définissez des points d’arrêt dans votre code, Office sur le Web risque de créer une erreur indiquant qu’il est impossible d’enregistrer.

## <a name="see-also"></a>Voir aussi

- [Meilleures pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md)
