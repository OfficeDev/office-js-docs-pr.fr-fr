---
title: Élément Supertip dans le fichier manifest
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: ab280ec550a58f85082c36a24f5f7c3b4112a214
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325233"
---
# <a name="supertip"></a><span data-ttu-id="27609-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="27609-102">Supertip</span></span>

<span data-ttu-id="27609-p101">Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="27609-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="27609-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="27609-105">Child elements</span></span>

|  <span data-ttu-id="27609-106">Élément</span><span class="sxs-lookup"><span data-stu-id="27609-106">Element</span></span> |  <span data-ttu-id="27609-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="27609-107">Required</span></span>  |  <span data-ttu-id="27609-108">Description</span><span class="sxs-lookup"><span data-stu-id="27609-108">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="27609-109">Titre</span><span class="sxs-lookup"><span data-stu-id="27609-109">Title</span></span>](#title) | <span data-ttu-id="27609-110">Oui</span><span class="sxs-lookup"><span data-stu-id="27609-110">Yes</span></span> | <span data-ttu-id="27609-111">Texte de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="27609-111">The text for the supertip.</span></span> |
| [<span data-ttu-id="27609-112">Description</span><span class="sxs-lookup"><span data-stu-id="27609-112">Description</span></span>](#description) | <span data-ttu-id="27609-113">Oui</span><span class="sxs-lookup"><span data-stu-id="27609-113">Yes</span></span> | <span data-ttu-id="27609-114">Description de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="27609-114">The description for the supertip.</span></span><br><span data-ttu-id="27609-115">**Remarque**: (Outlook) seuls les clients Windows et Mac sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="27609-115">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="27609-116">Titre</span><span class="sxs-lookup"><span data-stu-id="27609-116">Title</span></span>

<span data-ttu-id="27609-117">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="27609-117">Required.</span></span> <span data-ttu-id="27609-118">Texte de la propriété SuperTip.</span><span class="sxs-lookup"><span data-stu-id="27609-118">The text for the supertip.</span></span> <span data-ttu-id="27609-119">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="27609-119">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="27609-120">Description</span><span class="sxs-lookup"><span data-stu-id="27609-120">Description</span></span>

<span data-ttu-id="27609-121">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="27609-121">Required.</span></span> <span data-ttu-id="27609-122">Description de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="27609-122">The description for the supertip.</span></span> <span data-ttu-id="27609-123">L’attribut **RESID** doit être défini sur la valeur de l' **attribut ID** d’un élément **String** dans l’élément **LongStrings** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="27609-123">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="27609-124">Pour Outlook, seuls les clients Windows et Mac prennent en charge l’élément **Description** .</span><span class="sxs-lookup"><span data-stu-id="27609-124">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="27609-125">Exemple</span><span class="sxs-lookup"><span data-stu-id="27609-125">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
