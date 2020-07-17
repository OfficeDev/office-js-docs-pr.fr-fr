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
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a><span data-ttu-id="b63d8-103">Objets Window non pris en charge dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="b63d8-103">Window objects that are unsupported in Office Add-ins</span></span>

<span data-ttu-id="b63d8-104">Pour certaines versions de Windows et d’Office, les compléments s’exécutent dans le runtime Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="b63d8-104">For some versions of Windows and Office, add-ins run in an Internet Explorer 11 runtime.</span></span> <span data-ttu-id="b63d8-105">(Pour plus d’informations, consultez la rubrique [navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).) Certaines propriétés ou sous-propriétés de l' `window` objet global ne sont pas prises en charge dans Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="b63d8-105">(For details, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Some properties or subproperties of the global `window` object are not supported in Internet Explorer 11.</span></span> <span data-ttu-id="b63d8-106">Ces propriétés sont désactivées dans les compléments pour garantir que votre complément offre une expérience cohérente pour tous les utilisateurs, quel que soit le navigateur utilisé par le complément.</span><span class="sxs-lookup"><span data-stu-id="b63d8-106">These properties are disabled in add-ins to ensure that your add-in provides a consistent experience to all users, regardless of which browser the add-in is using.</span></span> <span data-ttu-id="b63d8-107">Cela permet également de charger correctement AngularJS.</span><span class="sxs-lookup"><span data-stu-id="b63d8-107">This also helps AngularJS load properly.</span></span>

<span data-ttu-id="b63d8-108">Voici une liste des propriétés désactivées.</span><span class="sxs-lookup"><span data-stu-id="b63d8-108">The following is a list of the disabled properties.</span></span> <span data-ttu-id="b63d8-109">La liste est une tâche en cours.</span><span class="sxs-lookup"><span data-stu-id="b63d8-109">The list is a work in progress.</span></span> <span data-ttu-id="b63d8-110">Si vous découvrez `window` d’autres propriétés qui ne fonctionnent pas dans des compléments, utilisez l’outil de commentaires ci-dessous pour nous les informer.</span><span class="sxs-lookup"><span data-stu-id="b63d8-110">If you discover additional `window` properties that do not work in add-ins, please use the feedback tool below to tell us.</span></span>

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a><span data-ttu-id="b63d8-111">Consultez également</span><span class="sxs-lookup"><span data-stu-id="b63d8-111">See also</span></span>

- [<span data-ttu-id="b63d8-112">Navigateurs utilisés par les compléments Office</span><span class="sxs-lookup"><span data-stu-id="b63d8-112">Browsers used by Office Add-ins</span></span>](../concepts/browsers-used-by-office-web-add-ins.md)