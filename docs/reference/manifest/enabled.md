---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de complément est désactivée au lancement du complément.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 2849689fec99190c3a9b039c6c04069bc8194ee1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611567"
---
# <a name="enabled-element"></a><span data-ttu-id="0a80a-103">Élément Enabled</span><span class="sxs-lookup"><span data-stu-id="0a80a-103">Enabled element</span></span>

<span data-ttu-id="0a80a-104">Indique si un [bouton](control.md#button-control) ou un contrôle de [menu](control.md#menu-dropdown-button-controls) est activé au lancement du complément.</span><span class="sxs-lookup"><span data-stu-id="0a80a-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="0a80a-105">L’élément **Enabled** est un élément enfant de [Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="0a80a-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="0a80a-106">Si ce paramètre est omis, la valeur par défaut est `true` .</span><span class="sxs-lookup"><span data-stu-id="0a80a-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="0a80a-107">Le contrôle parent peut également être activé et désactivé par programme.</span><span class="sxs-lookup"><span data-stu-id="0a80a-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="0a80a-108">Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="0a80a-108">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="0a80a-109">Exemple</span><span class="sxs-lookup"><span data-stu-id="0a80a-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```
