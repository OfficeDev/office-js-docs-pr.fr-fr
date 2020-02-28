---
title: Élément MobileFormFactor dans le fichier manifest
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 34106011cb855b6ac7c6d0fc21c16fd13e52b281
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324840"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="a8096-102">Élément MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="a8096-102">MobileFormFactor element</span></span>

<span data-ttu-id="a8096-p101">Spécifie les paramètres d’un complément pour le facteur de forme pour environnement mobile. Il contient toutes les informations de complément pour ce facteur de forme pour environnement mobile pour le nœud **Resources**.</span><span class="sxs-lookup"><span data-stu-id="a8096-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="a8096-105">Chaque définition **MobileFormFactor** contient l’élément **FunctionFile** et un ou plusieurs éléments **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="a8096-105">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="a8096-106">Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="a8096-106">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="a8096-p103">L’élément **MobileFormFactor** est défini dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) le contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="a8096-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="a8096-109">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="a8096-109">Child elements</span></span>

| <span data-ttu-id="a8096-110">Élément</span><span class="sxs-lookup"><span data-stu-id="a8096-110">Element</span></span>                               | <span data-ttu-id="a8096-111">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a8096-111">Required</span></span> | <span data-ttu-id="a8096-112">Description</span><span class="sxs-lookup"><span data-stu-id="a8096-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="a8096-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="a8096-113">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="a8096-114">Oui</span><span class="sxs-lookup"><span data-stu-id="a8096-114">Yes</span></span>      | <span data-ttu-id="a8096-115">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="a8096-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="a8096-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="a8096-116">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="a8096-117">Oui</span><span class="sxs-lookup"><span data-stu-id="a8096-117">Yes</span></span>      | <span data-ttu-id="a8096-118">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a8096-118">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="a8096-119">Exemple MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="a8096-119">MobileFormFactor example</span></span>

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
