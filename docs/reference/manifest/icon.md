---
title: Élément Icon dans le fichier manifeste
description: Définit les éléments Image pour les contrôles de bouton ou de menu.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ff16e6c0fbbf6d1c54508b4460ed3e02e899db03
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771332"
---
# <a name="icon-element"></a><span data-ttu-id="3405e-103">Icon, élément</span><span class="sxs-lookup"><span data-stu-id="3405e-103">Icon element</span></span>

<span data-ttu-id="3405e-104">Définit les éléments **Image** pour les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="3405e-104">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="3405e-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="3405e-105">Attributes</span></span>

|  <span data-ttu-id="3405e-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="3405e-106">Attribute</span></span>  |  <span data-ttu-id="3405e-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="3405e-107">Required</span></span>  |  <span data-ttu-id="3405e-108">Description</span><span class="sxs-lookup"><span data-stu-id="3405e-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3405e-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="3405e-109">**xsi:type**</span></span>  |  <span data-ttu-id="3405e-110">Non</span><span class="sxs-lookup"><span data-stu-id="3405e-110">No</span></span>  | <span data-ttu-id="3405e-p101">Type d’icône en cours de définition. Uniquement applicable aux icônes dans des facteurs de forme pour environnement mobile. Pour les éléments **Icon** contenus dans un élément [MobileFormFactor](mobileformfactor.md), cet attribut doit être défini sur `bt:MobileIconList`.</span><span class="sxs-lookup"><span data-stu-id="3405e-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="3405e-114">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="3405e-114">Child elements</span></span>

|  <span data-ttu-id="3405e-115">Élément</span><span class="sxs-lookup"><span data-stu-id="3405e-115">Element</span></span> |  <span data-ttu-id="3405e-116">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="3405e-116">Required</span></span>  |  <span data-ttu-id="3405e-117">Description</span><span class="sxs-lookup"><span data-stu-id="3405e-117">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3405e-118">Image</span><span class="sxs-lookup"><span data-stu-id="3405e-118">Image</span></span>](#image)        | <span data-ttu-id="3405e-119">Oui</span><span class="sxs-lookup"><span data-stu-id="3405e-119">Yes</span></span> |   <span data-ttu-id="3405e-120">Attribut resid d’une image à utiliser</span><span class="sxs-lookup"><span data-stu-id="3405e-120">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="3405e-121">Image</span><span class="sxs-lookup"><span data-stu-id="3405e-121">Image</span></span>

<span data-ttu-id="3405e-122">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="3405e-122">An image for the button.</span></span> <span data-ttu-id="3405e-123">L’attribut **RESID** ne peut pas contenir plus de 32 caractères et doit être défini sur la valeur de l’attribut **ID** d’un élément **image** dans l’élément **images** de l’élément [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="3405e-123">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element.</span></span> <span data-ttu-id="3405e-124">L’attribut **size** indique la taille de l’image en pixels.</span><span class="sxs-lookup"><span data-stu-id="3405e-124">The **size** attribute indicates the size in pixels of the image.</span></span> <span data-ttu-id="3405e-125">Trois tailles d’image sont requises (16, 32 et 80 pixels) et cinq autres tailles sont prises en charge (20, 24, 40, 48 et 64 pixels).|</span><span class="sxs-lookup"><span data-stu-id="3405e-125">Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="3405e-126">Configuration requise supplémentaire pour les facteurs de forme pour environnement mobile</span><span class="sxs-lookup"><span data-stu-id="3405e-126">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="3405e-p103">Lorsque l’élément **Icon** parent est un descendant de l’élément [MobileFormFactor](mobileformfactor.md), la taille minimale requise est légèrement différente. Le manifeste doit fournir au minimum les tailles 25, 32 et 48 pixels. Chaque taille fournie doit apparaître trois fois, avec un ensemble d’attributs `scale` défini sur `1`, `2` ou `3`.</span><span class="sxs-lookup"><span data-stu-id="3405e-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```
