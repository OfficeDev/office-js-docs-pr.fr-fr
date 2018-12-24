---
title: Élément Namespace dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 8000ea5774b38dd038888c686a33127a2d5bc482
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432325"
---
# <a name="namespace-element"></a><span data-ttu-id="099a2-102">Élément Namespace</span><span class="sxs-lookup"><span data-stu-id="099a2-102">Namespace element</span></span>

<span data-ttu-id="099a2-103">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="099a2-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="099a2-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="099a2-104">Attributes</span></span>

|  <span data-ttu-id="099a2-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="099a2-105">Attribute</span></span>  |  <span data-ttu-id="099a2-106">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="099a2-106">Required</span></span>  |  <span data-ttu-id="099a2-107">Description</span><span class="sxs-lookup"><span data-stu-id="099a2-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="099a2-108">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="099a2-108">**resid="namespace"**</span></span>  |  <span data-ttu-id="099a2-109">Oui</span><span class="sxs-lookup"><span data-stu-id="099a2-109">Yes</span></span>  | <span data-ttu-id="099a2-110">Doit correspondre à votre fonction personnalisée spécifiée dans le titre ShortStrings de l’élément[ressources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="099a2-110">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="099a2-111">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="099a2-111">Child elements</span></span>

<span data-ttu-id="099a2-112">Néant</span><span class="sxs-lookup"><span data-stu-id="099a2-112">None</span></span>

## <a name="example"></a><span data-ttu-id="099a2-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="099a2-113">Example</span></span>

```xml
<Namespace resid="namespace" />
```
