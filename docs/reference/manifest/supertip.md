---
title: Élément Supertip dans le fichier manifest
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdbba342fa591ddff3faf94ecd63a4740fb904da
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450540"
---
# <a name="supertip"></a><span data-ttu-id="290df-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="290df-102">Supertip</span></span>

<span data-ttu-id="290df-p101">Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="290df-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="290df-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="290df-105">Child elements</span></span>

|  <span data-ttu-id="290df-106">Élément</span><span class="sxs-lookup"><span data-stu-id="290df-106">Element</span></span> |  <span data-ttu-id="290df-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="290df-107">Required</span></span>  |  <span data-ttu-id="290df-108">Description</span><span class="sxs-lookup"><span data-stu-id="290df-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="290df-109">Titre</span><span class="sxs-lookup"><span data-stu-id="290df-109">Title</span></span>](#title)        | <span data-ttu-id="290df-110">Oui</span><span class="sxs-lookup"><span data-stu-id="290df-110">Yes</span></span> |   <span data-ttu-id="290df-111">Texte de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="290df-111">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="290df-112">Description</span><span class="sxs-lookup"><span data-stu-id="290df-112">Description</span></span>](#description)  | <span data-ttu-id="290df-113">Oui</span><span class="sxs-lookup"><span data-stu-id="290df-113">Yes</span></span> |  <span data-ttu-id="290df-114">Description de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="290df-114">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="290df-115">Titre</span><span class="sxs-lookup"><span data-stu-id="290df-115">Title</span></span>

<span data-ttu-id="290df-p102">Obligatoire. Texte de la propriété SuperTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="290df-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="290df-119">Description</span><span class="sxs-lookup"><span data-stu-id="290df-119">Description</span></span>

<span data-ttu-id="290df-p103">Obligatoire. Description de la propriété SuperTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **LongStrings** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="290df-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="290df-123">Exemple</span><span class="sxs-lookup"><span data-stu-id="290df-123">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
