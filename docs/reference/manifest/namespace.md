---
title: Élément Namespace dans le fichier manifest
description: L’élément namespace définit l’espace de noms qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 45fd0caa039fdeb885cba4b739750fbd8b642252
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718055"
---
# <a name="namespace-element"></a><span data-ttu-id="98794-103">Élément Namespace</span><span class="sxs-lookup"><span data-stu-id="98794-103">Namespace element</span></span>

<span data-ttu-id="98794-104">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="98794-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="98794-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="98794-105">Attributes</span></span>

|  <span data-ttu-id="98794-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="98794-106">Attribute</span></span>  |  <span data-ttu-id="98794-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="98794-107">Required</span></span>  |  <span data-ttu-id="98794-108">Description</span><span class="sxs-lookup"><span data-stu-id="98794-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="98794-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="98794-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="98794-110">Oui</span><span class="sxs-lookup"><span data-stu-id="98794-110">Yes</span></span>  | <span data-ttu-id="98794-111">Doit correspondre à votre fonction personnalisée spécifiée dans le titre ShortStrings de l’élément[ressources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="98794-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="98794-112">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="98794-112">Child elements</span></span>

<span data-ttu-id="98794-113">Aucun</span><span class="sxs-lookup"><span data-stu-id="98794-113">None</span></span>

## <a name="example"></a><span data-ttu-id="98794-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="98794-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
