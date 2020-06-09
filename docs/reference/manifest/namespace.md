---
title: Élément Namespace dans le fichier manifest
description: L’élément namespace définit l’espace de noms qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f4b3510c6c137bd303af8a3eaac8ebe66c5f4dc7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612233"
---
# <a name="namespace-element"></a><span data-ttu-id="4979d-103">Élément Namespace</span><span class="sxs-lookup"><span data-stu-id="4979d-103">Namespace element</span></span>

<span data-ttu-id="4979d-104">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="4979d-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="4979d-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="4979d-105">Attributes</span></span>

|  <span data-ttu-id="4979d-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="4979d-106">Attribute</span></span>  |  <span data-ttu-id="4979d-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="4979d-107">Required</span></span>  |  <span data-ttu-id="4979d-108">Description</span><span class="sxs-lookup"><span data-stu-id="4979d-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="4979d-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="4979d-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="4979d-110">Non</span><span class="sxs-lookup"><span data-stu-id="4979d-110">No</span></span>  | <span data-ttu-id="4979d-111">Doit correspondre à votre fonction personnalisée spécifiée dans le titre ShortStrings de l’élément[ressources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="4979d-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="4979d-112">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="4979d-112">Child elements</span></span>

<span data-ttu-id="4979d-113">Aucun</span><span class="sxs-lookup"><span data-stu-id="4979d-113">None</span></span>

## <a name="example"></a><span data-ttu-id="4979d-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="4979d-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
