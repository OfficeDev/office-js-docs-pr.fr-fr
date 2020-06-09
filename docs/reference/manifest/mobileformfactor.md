---
title: Élément MobileFormFactor dans le fichier manifest
description: L’élément MobileFormFactor spécifie les paramètres de facteur de forme mobile pour un complément.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 64a7681ca23becf42af1ba435aae4d509e6ad1ba
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612226"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="91891-103">Élément MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="91891-103">MobileFormFactor element</span></span>

<span data-ttu-id="91891-p101">Spécifie les paramètres d’un complément pour le facteur de forme pour environnement mobile. Il contient toutes les informations de complément pour ce facteur de forme pour environnement mobile pour le nœud **Resources**.</span><span class="sxs-lookup"><span data-stu-id="91891-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="91891-106">Chaque définition **MobileFormFactor** contient l’élément **FunctionFile** et un ou plusieurs éléments **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="91891-106">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="91891-107">Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="91891-107">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="91891-p103">L’élément **MobileFormFactor** est défini dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) le contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="91891-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="91891-110">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="91891-110">Child elements</span></span>

| <span data-ttu-id="91891-111">Élément</span><span class="sxs-lookup"><span data-stu-id="91891-111">Element</span></span>                               | <span data-ttu-id="91891-112">Requis</span><span class="sxs-lookup"><span data-stu-id="91891-112">Required</span></span> | <span data-ttu-id="91891-113">Description</span><span class="sxs-lookup"><span data-stu-id="91891-113">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="91891-114">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="91891-114">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="91891-115">Oui</span><span class="sxs-lookup"><span data-stu-id="91891-115">Yes</span></span>      | <span data-ttu-id="91891-116">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="91891-116">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="91891-117">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="91891-117">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="91891-118">Oui</span><span class="sxs-lookup"><span data-stu-id="91891-118">Yes</span></span>      | <span data-ttu-id="91891-119">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="91891-119">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="91891-120">Exemple MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="91891-120">MobileFormFactor example</span></span>

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
