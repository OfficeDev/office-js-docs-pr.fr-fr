---
title: Élément Icon dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 476e5720e4959c3c766a7ae6206f2bf12731bfd2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432269"
---
# <a name="icon-element"></a><span data-ttu-id="fbf38-102">Icon, élément</span><span class="sxs-lookup"><span data-stu-id="fbf38-102">Icon element</span></span>

<span data-ttu-id="fbf38-103">Définit les éléments **Image** pour les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="fbf38-103">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="fbf38-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="fbf38-104">Attributes</span></span>

|  <span data-ttu-id="fbf38-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="fbf38-105">Attribute</span></span>  |  <span data-ttu-id="fbf38-106">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="fbf38-106">Required</span></span>  |  <span data-ttu-id="fbf38-107">Description</span><span class="sxs-lookup"><span data-stu-id="fbf38-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="fbf38-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="fbf38-108">**xsi:type**</span></span>  |  <span data-ttu-id="fbf38-109">Non</span><span class="sxs-lookup"><span data-stu-id="fbf38-109">No</span></span>  | <span data-ttu-id="fbf38-p101">Type d’icône en cours de définition. Uniquement applicable aux icônes dans des facteurs de forme pour environnement mobile. Pour les éléments **Icon** contenus dans un élément [MobileFormFactor](mobileformfactor.md), cet attribut doit être défini sur `bt:MobileIconList`.</span><span class="sxs-lookup"><span data-stu-id="fbf38-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="fbf38-113">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="fbf38-113">Child elements</span></span>

|  <span data-ttu-id="fbf38-114">Élément</span><span class="sxs-lookup"><span data-stu-id="fbf38-114">Element</span></span> |  <span data-ttu-id="fbf38-115">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="fbf38-115">Required</span></span>  |  <span data-ttu-id="fbf38-116">Description</span><span class="sxs-lookup"><span data-stu-id="fbf38-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fbf38-117">Image</span><span class="sxs-lookup"><span data-stu-id="fbf38-117">Image</span></span>](#image)        | <span data-ttu-id="fbf38-118">Oui</span><span class="sxs-lookup"><span data-stu-id="fbf38-118">Yes</span></span> |   <span data-ttu-id="fbf38-119">Attribut resid d’une image à utiliser</span><span class="sxs-lookup"><span data-stu-id="fbf38-119">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="fbf38-120">Image</span><span class="sxs-lookup"><span data-stu-id="fbf38-120">Image</span></span>

<span data-ttu-id="fbf38-p102">Image du bouton. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image** dans l’élément **Images** dans l’élément [Resources](resources.md). L’attribut **size** indique la taille de l’image en pixels. Trois tailles d’image sont requises (16, 32 et 80 pixels) et cinq autres tailles sont prises en charge (20, 24, 40, 48 et 64 pixels).|</span><span class="sxs-lookup"><span data-stu-id="fbf38-p102">An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="fbf38-125">Configuration requise supplémentaire pour les facteurs de forme pour environnement mobile</span><span class="sxs-lookup"><span data-stu-id="fbf38-125">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="fbf38-p103">Lorsque l’élément **Icon** parent est un descendant de l’élément [MobileFormFactor](mobileformfactor.md), la taille minimale requise est légèrement différente. Le manifeste doit fournir au minimum les tailles 25, 32 et 48 pixels. Chaque taille fournie doit apparaître trois fois, avec un ensemble d’attributs `scale` défini sur `1`, `2` ou `3`.</span><span class="sxs-lookup"><span data-stu-id="fbf38-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

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