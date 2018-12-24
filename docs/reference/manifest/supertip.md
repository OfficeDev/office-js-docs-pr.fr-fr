---
title: Élément Supertip dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: bae997eda8e1055c5be76382456ba83acca7b91c
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433669"
---
# <a name="supertip"></a><span data-ttu-id="31756-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="31756-102">Supertip</span></span>

<span data-ttu-id="31756-p101">Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="31756-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="31756-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="31756-105">Child elements</span></span>

|  <span data-ttu-id="31756-106">Élément</span><span class="sxs-lookup"><span data-stu-id="31756-106">Element</span></span> |  <span data-ttu-id="31756-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="31756-107">Required</span></span>  |  <span data-ttu-id="31756-108">Description</span><span class="sxs-lookup"><span data-stu-id="31756-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="31756-109">Title</span><span class="sxs-lookup"><span data-stu-id="31756-109">Title</span></span>](#title)        | <span data-ttu-id="31756-110">Oui</span><span class="sxs-lookup"><span data-stu-id="31756-110">Yes</span></span> |   <span data-ttu-id="31756-111">Texte de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="31756-111">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="31756-112">Description</span><span class="sxs-lookup"><span data-stu-id="31756-112">Description</span></span>](#description)  | <span data-ttu-id="31756-113">Oui</span><span class="sxs-lookup"><span data-stu-id="31756-113">Yes</span></span> |  <span data-ttu-id="31756-114">Description de l’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="31756-114">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="31756-115">Titre</span><span class="sxs-lookup"><span data-stu-id="31756-115">Title</span></span>

<span data-ttu-id="31756-p102">Obligatoire. Texte de la propriété SuperTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="31756-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="31756-119">Description</span><span class="sxs-lookup"><span data-stu-id="31756-119">Description</span></span>

<span data-ttu-id="31756-p103">Obligatoire. Description de la propriété SuperTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **LongStrings** dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="31756-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="31756-123">Exemple</span><span class="sxs-lookup"><span data-stu-id="31756-123">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
