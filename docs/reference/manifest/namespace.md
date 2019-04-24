---
title: Élément Namespace dans le fichier manifest
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: faf77fe8b6bddc734f1b47eb544ffe7e1e7c4aaa
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452101"
---
# <a name="namespace-element"></a><span data-ttu-id="5bc92-102">Élément Namespace</span><span class="sxs-lookup"><span data-stu-id="5bc92-102">Namespace element</span></span>

<span data-ttu-id="5bc92-103">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="5bc92-103">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5bc92-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="5bc92-104">Attributes</span></span>

|  <span data-ttu-id="5bc92-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="5bc92-105">Attribute</span></span>  |  <span data-ttu-id="5bc92-106">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="5bc92-106">Required</span></span>  |  <span data-ttu-id="5bc92-107">Description</span><span class="sxs-lookup"><span data-stu-id="5bc92-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5bc92-108">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="5bc92-108">**resid="namespace"**</span></span>  |  <span data-ttu-id="5bc92-109">Oui</span><span class="sxs-lookup"><span data-stu-id="5bc92-109">Yes</span></span>  | <span data-ttu-id="5bc92-110">Doit correspondre à votre fonction personnalisée spécifiée dans le titre ShortStrings de l’élément[ressources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="5bc92-110">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="5bc92-111">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="5bc92-111">Child elements</span></span>

<span data-ttu-id="5bc92-112">Aucun</span><span class="sxs-lookup"><span data-stu-id="5bc92-112">None</span></span>

## <a name="example"></a><span data-ttu-id="5bc92-113">Exemple</span><span class="sxs-lookup"><span data-stu-id="5bc92-113">Example</span></span>

```xml
<Namespace resid="namespace" />
```
