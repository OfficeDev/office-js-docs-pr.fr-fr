---
title: Élément Namespace dans le fichier manifest
description: L’élément namespace définit l’espace de noms qu’une fonction personnalisée utilise dans Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 342f5ebcafa861838956f1033f8597cf05e60215
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771257"
---
# <a name="namespace-element"></a><span data-ttu-id="d1cb6-103">Élément Namespace</span><span class="sxs-lookup"><span data-stu-id="d1cb6-103">Namespace element</span></span>

<span data-ttu-id="d1cb6-104">Définit les paramètres de script utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="d1cb6-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="d1cb6-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="d1cb6-105">Attributes</span></span>

|  <span data-ttu-id="d1cb6-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="d1cb6-106">Attribute</span></span>  |  <span data-ttu-id="d1cb6-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="d1cb6-107">Required</span></span>  |  <span data-ttu-id="d1cb6-108">Description</span><span class="sxs-lookup"><span data-stu-id="d1cb6-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d1cb6-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="d1cb6-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="d1cb6-110">Non</span><span class="sxs-lookup"><span data-stu-id="d1cb6-110">No</span></span>  | <span data-ttu-id="d1cb6-111">Doit correspondre à votre fonction personnalisée spécifiée dans le titre ShortStrings de l’élément[ressources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="d1cb6-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> <span data-ttu-id="d1cb6-112">Il ne peut pas comporter plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="d1cb6-112">Can be no more than 32 characters.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="d1cb6-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="d1cb6-113">Child elements</span></span>

<span data-ttu-id="d1cb6-114">Aucun</span><span class="sxs-lookup"><span data-stu-id="d1cb6-114">None</span></span>

## <a name="example"></a><span data-ttu-id="d1cb6-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="d1cb6-115">Example</span></span>

```xml
<Namespace resid="namespace" />
```
