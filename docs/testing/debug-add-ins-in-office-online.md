---
title: Débogage de compléments dans Office sur le web
description: Découvrez comment utiliser Office sur le web pour tester et déboguer vos compléments.
ms.date: 03/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: b365be937058f818a97dd7a73176a56f76b36098
ms.sourcegitcommit: a32f5613d2bb44a8c812d7d407f106422a530f7a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/14/2022
ms.locfileid: "67674624"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Débogage de compléments dans Office sur le web

Cet article explique comment utiliser Office sur le Web pour déboguer vos compléments. Utilisez cette technique :

- Pour déboguer des compléments sur un ordinateur qui n’exécute pas Windows ou le client&mdash;de bureau Office, par exemple, si vous développez sur un Mac ou Linux.
- Comme autre processus de débogage si vous ne pouvez pas ou ne souhaitez pas déboguer dans un IDE, tel que Visual Studio ou Visual Studio Code.

Cet article part du principe que vous disposez d’un projet de complément qui doit être débogué. Si vous souhaitez simplement vous exercer au débogage sur le web, créez un projet à l’aide de l’un des démarrages rapides pour des applications Office spécifiques, comme ce [guide de démarrage rapide pour Word](../quickstarts/word-quickstart.md).

## <a name="debug-your-add-in"></a>Déboguer votre complément

Pour déboguer votre complément à l’aide d’Office sur le web, procédez comme suit :

1. Exécutez le projet sur localhost et chargez-le sur un document dans Office sur le Web. Pour obtenir des instructions détaillées sur le chargement indépendant, consultez [Chargement indépendant des compléments Office sur le web](sideload-office-add-ins-for-testing.md#manually-sideload-an-add-in-to-office-on-the-web).

2. Ouvrez les outils de développement du navigateur. Pour ce faire, appuyez généralement sur F12. Ouvrez l’outil débogueur et utilisez-le pour définir des points d’arrêt et observer des variables. Pour obtenir de l’aide détaillée sur l’utilisation de l’outil de votre navigateur, consultez l’une des rubriques suivantes :

   - [Firefox](https://firefox-source-docs.mozilla.org/devtools-user/index.html)
   - [Safari](https://support.apple.com/guide/safari/use-the-developer-tools-in-the-develop-menu-sfri20948/mac)
   - [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
   - [Déboguer des compléments à l’aide des outils de développement pour la version héritée Edge](debug-add-ins-using-devtools-edge-legacy.md)

   > [!NOTE]
   > Office sur le Web ne s’ouvre pas dans Internet Explorer.

## <a name="potential-issues"></a>Problèmes potentiels

Voici quelques problèmes que vous pouvez rencontrer lors du débogage.

- Certaines erreurs JavaScript peuvent provenir d’Office sur le web.

- Le navigateur peut afficher une erreur relative à un certificat non valide que vous devrez contourner. Le processus d’exécution de cette opération varie en fonction du navigateur et des interfaces utilisateur des différents navigateurs permettant d’effectuer cette modification régulièrement. Vous devez effectuer une recherche dans l’aide du navigateur ou rechercher des instructions en ligne. (Par exemple, recherchez « Avertissement de certificat Microsoft Edge non valide ».) La plupart des navigateurs, sur la page d’avertissement, comportent un lien qui vous permet d’accéder à la page du complément. Par exemple, Microsoft Edge comporte un lien « Accéder à la page web (non recommandé) ». En général, vous devez passer par ce lien chaque fois que le complément est rechargé. Pour un contournement plus long, consultez l’aide comme suggéré.

- Si vous définissez des points d’arrêt dans votre code, Office sur le Web pouvez générer une erreur indiquant qu’il n’est pas en mesure d’enregistrer.

## <a name="see-also"></a>Voir aussi

- [Meilleures pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md)
