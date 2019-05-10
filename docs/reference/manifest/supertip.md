---
title: Élément Supertip dans le fichier manifest
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 269a3723db6f98cdb25c61e5a88608c5fb5f3191
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659654"
---
# <a name="supertip"></a><span data-ttu-id="79624-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="79624-102">Supertip</span></span>

<span data-ttu-id="79624-p101">Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="79624-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="79624-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="79624-105">Child elements</span></span>

|  <span data-ttu-id="79624-106">Élément</span><span class="sxs-lookup"><span data-stu-id="79624-106">Element</span></span> |  <span data-ttu-id="79624-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="79624-107">Required</span></span>  |  <span data-ttu-id="79624-108">Description</span><span class="sxs-lookup"><span data-stu-id="79624-108">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="79624-109">Titre</span><span class="sxs-lookup"><span data-stu-id="79624-109">Title</span></span>](#title) | <span data-ttu-id="79624-110">Oui</span><span class="sxs-lookup"><span data-stu-id="79624-110">Yes</span></span> | <span data-ttu-id="79624-111">Texte de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="79624-111">The text for the supertip.</span></span> |
| [<span data-ttu-id="79624-112">Description</span><span class="sxs-lookup"><span data-stu-id="79624-112">Description</span></span>](#description) | <span data-ttu-id="79624-113">Oui</span><span class="sxs-lookup"><span data-stu-id="79624-113">Yes</span></span> | <span data-ttu-id="79624-114">Description de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="79624-114">The description for the supertip.</span></span><br><span data-ttu-id="79624-115">**Remarque**: (Outlook) seuls les clients Windows et Mac sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="79624-115">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="79624-116">Titre</span><span class="sxs-lookup"><span data-stu-id="79624-116">Title</span></span>

<span data-ttu-id="79624-p102">Obligatoire. Texte de la propriété SuperTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="79624-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="79624-120">Description</span><span class="sxs-lookup"><span data-stu-id="79624-120">Description</span></span>

<span data-ttu-id="79624-p103">Obligatoire. Description de la propriété SuperTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **LongStrings** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="79624-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="79624-124">Pour Outlook, seuls les clients Windows et Mac prennent en charge l’élément **Description** .</span><span class="sxs-lookup"><span data-stu-id="79624-124">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="79624-125">Exemple</span><span class="sxs-lookup"><span data-stu-id="79624-125">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
