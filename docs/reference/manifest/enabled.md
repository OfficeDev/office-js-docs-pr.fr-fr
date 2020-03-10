---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de complément est désactivée au lancement du complément.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: a47ab97ff5a159c73bea52f130ce0c16efe2b6b6
ms.sourcegitcommit: 0e7ed44019d6564c79113639af831ea512fa0a13
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/09/2020
ms.locfileid: "42566202"
---
# <a name="enabled-element"></a><span data-ttu-id="2e4a6-103">Élément Enabled</span><span class="sxs-lookup"><span data-stu-id="2e4a6-103">Enabled element</span></span>

<span data-ttu-id="2e4a6-104">Indique si un [bouton](control.md#button-control) ou un contrôle de [menu](control.md#menu-dropdown-button-controls) est activé au lancement du complément.</span><span class="sxs-lookup"><span data-stu-id="2e4a6-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="2e4a6-105">L’élément **Enabled** est un élément enfant de [Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="2e4a6-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="2e4a6-106">Si ce paramètre est omis, la valeur par `true`défaut est.</span><span class="sxs-lookup"><span data-stu-id="2e4a6-106">If it is omitted, the default is `true`.</span></span> 

<span data-ttu-id="2e4a6-107">Le contrôle parent peut également être activé et désactivé par programme.</span><span class="sxs-lookup"><span data-stu-id="2e4a6-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="2e4a6-108">Pour plus d’informations, consultez la rubrique [activer et désactiver les commandes de complément](/office/dev/add-ins/design/disable-add-in-commands).</span><span class="sxs-lookup"><span data-stu-id="2e4a6-108">For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).</span></span>

## <a name="example"></a><span data-ttu-id="2e4a6-109">Exemple</span><span class="sxs-lookup"><span data-stu-id="2e4a6-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```

