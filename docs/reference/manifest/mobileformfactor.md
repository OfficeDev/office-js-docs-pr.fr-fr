---
title: Élément MobileFormFactor dans le fichier manifest
description: L’élément MobileFormFactor spécifie les paramètres de facteur de forme mobile pour un complément.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 954fff5d1e701ce53a6ad82fa276c048ca6d6f3a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720588"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="b1797-103">Élément MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="b1797-103">MobileFormFactor element</span></span>

<span data-ttu-id="b1797-p101">Spécifie les paramètres d’un complément pour le facteur de forme pour environnement mobile. Il contient toutes les informations de complément pour ce facteur de forme pour environnement mobile pour le nœud **Resources**.</span><span class="sxs-lookup"><span data-stu-id="b1797-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="b1797-106">Chaque définition **MobileFormFactor** contient l’élément **FunctionFile** et un ou plusieurs éléments **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="b1797-106">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="b1797-107">Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="b1797-107">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="b1797-p103">L’élément **MobileFormFactor** est défini dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) le contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="b1797-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b1797-110">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b1797-110">Child elements</span></span>

| <span data-ttu-id="b1797-111">Élément</span><span class="sxs-lookup"><span data-stu-id="b1797-111">Element</span></span>                               | <span data-ttu-id="b1797-112">Requis</span><span class="sxs-lookup"><span data-stu-id="b1797-112">Required</span></span> | <span data-ttu-id="b1797-113">Description</span><span class="sxs-lookup"><span data-stu-id="b1797-113">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="b1797-114">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="b1797-114">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="b1797-115">Oui</span><span class="sxs-lookup"><span data-stu-id="b1797-115">Yes</span></span>      | <span data-ttu-id="b1797-116">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="b1797-116">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="b1797-117">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="b1797-117">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="b1797-118">Oui</span><span class="sxs-lookup"><span data-stu-id="b1797-118">Yes</span></span>      | <span data-ttu-id="b1797-119">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b1797-119">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="b1797-120">Exemple MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="b1797-120">MobileFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
