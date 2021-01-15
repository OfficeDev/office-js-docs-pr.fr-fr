---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de complément est désactivée au lancement du complément.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771388"
---
# <a name="enabled-element"></a><span data-ttu-id="8dfa6-103">Élément Enabled</span><span class="sxs-lookup"><span data-stu-id="8dfa6-103">Enabled element</span></span>

<span data-ttu-id="8dfa6-104">Indique si un [bouton](control.md#button-control) ou un contrôle de [menu](control.md#menu-dropdown-button-controls) est activé au lancement du complément.</span><span class="sxs-lookup"><span data-stu-id="8dfa6-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="8dfa6-105">L’élément **Enabled** est un élément enfant de [Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="8dfa6-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="8dfa6-106">Si ce paramètre est omis, la valeur par défaut est `true` .</span><span class="sxs-lookup"><span data-stu-id="8dfa6-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="8dfa6-107">Cet élément est valide uniquement dans Excel ; autrement dit, lorsque l' `Name` attribut de l’élément [hôte](host.md) est « classeur ».</span><span class="sxs-lookup"><span data-stu-id="8dfa6-107">This element is only valid in Excel; that is, when the `Name` attribute of the [Host](host.md) element is "Workbook".</span></span>

<span data-ttu-id="8dfa6-108">Le contrôle parent peut également être activé et désactivé par programme.</span><span class="sxs-lookup"><span data-stu-id="8dfa6-108">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="8dfa6-109">Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="8dfa6-109">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="8dfa6-110">Exemple</span><span class="sxs-lookup"><span data-stu-id="8dfa6-110">Example</span></span>

```xml
<Enabled>false</Enabled>
```
