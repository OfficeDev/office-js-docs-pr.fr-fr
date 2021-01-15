---
title: Élément Supertip dans le fichier manifest
description: L’élément SuperTip définit une info-bulle riche (titre et Description).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 5e8b3850d99f6791726b1b2f0545c5fb4b52c554
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771297"
---
# <a name="supertip"></a><span data-ttu-id="75279-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="75279-103">Supertip</span></span>

<span data-ttu-id="75279-p101">Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="75279-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="75279-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="75279-106">Child elements</span></span>

|  <span data-ttu-id="75279-107">Élément</span><span class="sxs-lookup"><span data-stu-id="75279-107">Element</span></span> |  <span data-ttu-id="75279-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="75279-108">Required</span></span>  |  <span data-ttu-id="75279-109">Description</span><span class="sxs-lookup"><span data-stu-id="75279-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="75279-110">Titre</span><span class="sxs-lookup"><span data-stu-id="75279-110">Title</span></span>](#title) | <span data-ttu-id="75279-111">Oui</span><span class="sxs-lookup"><span data-stu-id="75279-111">Yes</span></span> | <span data-ttu-id="75279-112">Texte de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="75279-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="75279-113">Description</span><span class="sxs-lookup"><span data-stu-id="75279-113">Description</span></span>](#description) | <span data-ttu-id="75279-114">Oui</span><span class="sxs-lookup"><span data-stu-id="75279-114">Yes</span></span> | <span data-ttu-id="75279-115">Description de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="75279-115">The description for the supertip.</span></span><br><span data-ttu-id="75279-116">**Remarque**: (Outlook) seuls les clients Windows et Mac sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="75279-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="75279-117">Titre</span><span class="sxs-lookup"><span data-stu-id="75279-117">Title</span></span>

<span data-ttu-id="75279-118">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="75279-118">Required.</span></span> <span data-ttu-id="75279-119">Texte de la propriété SuperTip.</span><span class="sxs-lookup"><span data-stu-id="75279-119">The text for the supertip.</span></span> <span data-ttu-id="75279-120">L’attribut **RESID** ne peut pas contenir plus de 32 caractères et doit être défini sur la valeur de l’attribut **ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="75279-120">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="75279-121">Description</span><span class="sxs-lookup"><span data-stu-id="75279-121">Description</span></span>

<span data-ttu-id="75279-122">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="75279-122">Required.</span></span> <span data-ttu-id="75279-123">Description de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="75279-123">The description for the supertip.</span></span> <span data-ttu-id="75279-124">L’attribut **RESID** ne peut pas contenir plus de 32 caractères et doit être défini sur la valeur de l’attribut **ID** d’un élément **String** dans l’élément **LongStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="75279-124">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="75279-125">Pour Outlook, seuls les clients Windows et Mac prennent en charge l’élément **Description** .</span><span class="sxs-lookup"><span data-stu-id="75279-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="75279-126">Exemple</span><span class="sxs-lookup"><span data-stu-id="75279-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
