---
title: Élément Namespace dans le fichier manifest
description: L’élément namespace définit l’espace de noms qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eabd73d3be98271c81723787dd3d1bdb6ee2ebcd
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978668"
---
# <a name="namespace-element"></a><span data-ttu-id="d8fbf-103">Élément Namespace</span><span class="sxs-lookup"><span data-stu-id="d8fbf-103">Namespace element</span></span>

<span data-ttu-id="d8fbf-104">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="d8fbf-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="d8fbf-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="d8fbf-105">Attributes</span></span>

|  <span data-ttu-id="d8fbf-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="d8fbf-106">Attribute</span></span>  |  <span data-ttu-id="d8fbf-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="d8fbf-107">Required</span></span>  |  <span data-ttu-id="d8fbf-108">Description</span><span class="sxs-lookup"><span data-stu-id="d8fbf-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d8fbf-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="d8fbf-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="d8fbf-110">Non</span><span class="sxs-lookup"><span data-stu-id="d8fbf-110">No</span></span>  | <span data-ttu-id="d8fbf-111">Doit correspondre à votre fonction personnalisée spécifiée dans le titre ShortStrings de l’élément[ressources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="d8fbf-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="d8fbf-112">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="d8fbf-112">Child elements</span></span>

<span data-ttu-id="d8fbf-113">Aucun</span><span class="sxs-lookup"><span data-stu-id="d8fbf-113">None</span></span>

## <a name="example"></a><span data-ttu-id="d8fbf-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="d8fbf-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
