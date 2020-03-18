---
title: Élément Supertip dans le fichier manifest
description: L’élément SuperTip définit une info-bulle riche (titre et Description).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: cf88473b72979c839e5d55f44938fda19be24084
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720350"
---
# <a name="supertip"></a><span data-ttu-id="c1ab3-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="c1ab3-103">Supertip</span></span>

<span data-ttu-id="c1ab3-p101">Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="c1ab3-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c1ab3-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="c1ab3-106">Child elements</span></span>

|  <span data-ttu-id="c1ab3-107">Élément</span><span class="sxs-lookup"><span data-stu-id="c1ab3-107">Element</span></span> |  <span data-ttu-id="c1ab3-108">Requis</span><span class="sxs-lookup"><span data-stu-id="c1ab3-108">Required</span></span>  |  <span data-ttu-id="c1ab3-109">Description</span><span class="sxs-lookup"><span data-stu-id="c1ab3-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="c1ab3-110">Titre</span><span class="sxs-lookup"><span data-stu-id="c1ab3-110">Title</span></span>](#title) | <span data-ttu-id="c1ab3-111">Oui</span><span class="sxs-lookup"><span data-stu-id="c1ab3-111">Yes</span></span> | <span data-ttu-id="c1ab3-112">Texte de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="c1ab3-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="c1ab3-113">Description</span><span class="sxs-lookup"><span data-stu-id="c1ab3-113">Description</span></span>](#description) | <span data-ttu-id="c1ab3-114">Oui</span><span class="sxs-lookup"><span data-stu-id="c1ab3-114">Yes</span></span> | <span data-ttu-id="c1ab3-115">Description de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="c1ab3-115">The description for the supertip.</span></span><br><span data-ttu-id="c1ab3-116">**Remarque**: (Outlook) seuls les clients Windows et Mac sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="c1ab3-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="c1ab3-117">Titre</span><span class="sxs-lookup"><span data-stu-id="c1ab3-117">Title</span></span>

<span data-ttu-id="c1ab3-118">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="c1ab3-118">Required.</span></span> <span data-ttu-id="c1ab3-119">Texte de la propriété SuperTip.</span><span class="sxs-lookup"><span data-stu-id="c1ab3-119">The text for the supertip.</span></span> <span data-ttu-id="c1ab3-120">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="c1ab3-120">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="c1ab3-121">Description</span><span class="sxs-lookup"><span data-stu-id="c1ab3-121">Description</span></span>

<span data-ttu-id="c1ab3-122">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="c1ab3-122">Required.</span></span> <span data-ttu-id="c1ab3-123">Description de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="c1ab3-123">The description for the supertip.</span></span> <span data-ttu-id="c1ab3-124">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **LongStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="c1ab3-124">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="c1ab3-125">Pour Outlook, seuls les clients Windows et Mac prennent en charge l’élément **Description** .</span><span class="sxs-lookup"><span data-stu-id="c1ab3-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="c1ab3-126">Exemple</span><span class="sxs-lookup"><span data-stu-id="c1ab3-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
