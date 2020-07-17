---
title: Objets Window non pris en charge dans les compléments Office
description: Cet article spécifie certains objets d’exécution de fenêtre qui ne fonctionnent pas dans les compléments Office.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: d2560748841bd1e2a7708b25a8e51133563d1534
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160502"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>Objets Window non pris en charge dans les compléments Office

Pour certaines versions de Windows et d’Office, les compléments s’exécutent dans le runtime Internet Explorer 11. (Pour plus d’informations, consultez la rubrique [navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).) Certaines propriétés ou sous-propriétés de l' `window` objet global ne sont pas prises en charge dans Internet Explorer 11. Ces propriétés sont désactivées dans les compléments pour garantir que votre complément offre une expérience cohérente pour tous les utilisateurs, quel que soit le navigateur utilisé par le complément. Cela permet également de charger correctement AngularJS.

Voici une liste des propriétés désactivées. La liste est une tâche en cours. Si vous découvrez `window` d’autres propriétés qui ne fonctionnent pas dans des compléments, utilisez l’outil de commentaires ci-dessous pour nous les informer.

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>Consultez également

- [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md)