---
title: Objets Window non pris en Office des modules
description: Cet article spécifie certains des objets runtime de fenêtre qui ne fonctionnent pas dans les Office de fenêtre.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: d2560748841bd1e2a7708b25a8e51133563d1534
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939186"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>Objets Window non pris en Office des modules

Pour certaines versions de Windows et Office, les modules sont exécutés dans un runtime Internet Explorer 11. (Pour plus d’informations, voir [Browsers used by Office Add-ins.)](../concepts/browsers-used-by-office-web-add-ins.md) Certaines propriétés ou sous-propriétés de l’objet global ne sont pas pris en `window` charge dans Internet Explorer 11. Ces propriétés sont désactivées dans les add-ins pour garantir une expérience cohérente à tous les utilisateurs, quel que soit le navigateur utilisé par le add-in. Cela permet également à AngularJS de se charger correctement.

Voici une liste des propriétés désactivées. La liste est un travail en cours. Si vous découvrez des propriétés supplémentaires qui ne fonctionnent pas dans les compléments, utilisez l’outil de commentaires `window` ci-dessous pour nous en faire part.

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>Voir aussi

- [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md)