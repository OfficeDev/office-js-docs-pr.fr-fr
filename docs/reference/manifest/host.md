---
title: Élément Host dans le fichier manifeste
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: debb4d59f75ce974ffb21d853c6b65a579c4e685
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127568"
---
# <a name="host-element"></a><span data-ttu-id="25499-102">Élément Host</span><span class="sxs-lookup"><span data-stu-id="25499-102">Host element</span></span>

<span data-ttu-id="25499-103">Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.</span><span class="sxs-lookup"><span data-stu-id="25499-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="25499-104">La syntaxe des éléments **Host** varie selon que l’élément est défini dans le [manifeste de base](#basic-manifest) ou le nœud [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="25499-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="25499-105">Toutefois, la fonctionnalité est identique.</span><span class="sxs-lookup"><span data-stu-id="25499-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="25499-106">Manifeste de base</span><span class="sxs-lookup"><span data-stu-id="25499-106">Basic manifest</span></span>

<span data-ttu-id="25499-107">Lorsqu’il est défini dans le manifeste base (sous [OfficeApp](officeapp.md)), le type d’hôte est déterminé par l’attribut `Name`.</span><span class="sxs-lookup"><span data-stu-id="25499-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="25499-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="25499-108">Attributes</span></span>

| <span data-ttu-id="25499-109">Attribut</span><span class="sxs-lookup"><span data-stu-id="25499-109">Attribute</span></span>     | <span data-ttu-id="25499-110">Type</span><span class="sxs-lookup"><span data-stu-id="25499-110">Type</span></span>   | <span data-ttu-id="25499-111">Requis</span><span class="sxs-lookup"><span data-stu-id="25499-111">Required</span></span> | <span data-ttu-id="25499-112">Description</span><span class="sxs-lookup"><span data-stu-id="25499-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="25499-113">Name</span><span class="sxs-lookup"><span data-stu-id="25499-113">Name</span></span>](#name) | <span data-ttu-id="25499-114">string</span><span class="sxs-lookup"><span data-stu-id="25499-114">string</span></span> | <span data-ttu-id="25499-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="25499-115">required</span></span> | <span data-ttu-id="25499-116">Nom du type d’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="25499-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="25499-117">Nom</span><span class="sxs-lookup"><span data-stu-id="25499-117">Name</span></span>
<span data-ttu-id="25499-p102">Spécifie le type d’hôte ciblé par ce complément. La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="25499-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="25499-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="25499-120">`Document` (Word)</span></span>
- <span data-ttu-id="25499-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="25499-121">`Database` (Access)</span></span>
- <span data-ttu-id="25499-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="25499-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="25499-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="25499-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="25499-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="25499-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="25499-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="25499-125">`Project` (Project)</span></span>
- <span data-ttu-id="25499-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="25499-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="25499-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="25499-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="25499-128">Nœud VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="25499-128">VersionOverrides node</span></span>
<span data-ttu-id="25499-129">Lorsqu’il est défini dans [VersionOverrides](versionoverrides.md), le type d’hôte est déterminé par l’attribut `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="25499-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="25499-130">Attributs</span><span class="sxs-lookup"><span data-stu-id="25499-130">Attributes</span></span>

|  <span data-ttu-id="25499-131">Attribut</span><span class="sxs-lookup"><span data-stu-id="25499-131">Attribute</span></span>  |  <span data-ttu-id="25499-132">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="25499-132">Required</span></span>  |  <span data-ttu-id="25499-133">Description</span><span class="sxs-lookup"><span data-stu-id="25499-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="25499-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="25499-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="25499-135">Oui</span><span class="sxs-lookup"><span data-stu-id="25499-135">Yes</span></span>  | <span data-ttu-id="25499-136">Décrit l’hôte d’Office dans lequel ces paramètres s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="25499-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="25499-137">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="25499-137">Child elements</span></span>

|  <span data-ttu-id="25499-138">Élément</span><span class="sxs-lookup"><span data-stu-id="25499-138">Element</span></span> |  <span data-ttu-id="25499-139">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="25499-139">Required</span></span>  |  <span data-ttu-id="25499-140">Description</span><span class="sxs-lookup"><span data-stu-id="25499-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="25499-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="25499-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="25499-142">Oui</span><span class="sxs-lookup"><span data-stu-id="25499-142">Yes</span></span>   |  <span data-ttu-id="25499-143">Définit les paramètres pour le facteur de forme pour bureau.</span><span class="sxs-lookup"><span data-stu-id="25499-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="25499-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="25499-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="25499-145">Non</span><span class="sxs-lookup"><span data-stu-id="25499-145">No</span></span>   |  <span data-ttu-id="25499-146">Définit les paramètres pour le facteur de forme pour environnement mobile.</span><span class="sxs-lookup"><span data-stu-id="25499-146">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="25499-147">**Remarque:** Cet élément est pris en charge uniquement dans Outlook sur iOS.</span><span class="sxs-lookup"><span data-stu-id="25499-147">**Note:** This element is only supported in Outlook on iOS.</span></span> |
|  [<span data-ttu-id="25499-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="25499-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="25499-149">Non</span><span class="sxs-lookup"><span data-stu-id="25499-149">No</span></span>   |  <span data-ttu-id="25499-150">Définit les paramètres de tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="25499-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="25499-151">Utilisé uniquement par des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="25499-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="25499-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="25499-152">xsi:type</span></span>

<span data-ttu-id="25499-153">Contrôle à quel hôte Office (Word, Excel, PowerPoint, Outlook, OneNote) s’applique également les paramètres contenus.</span><span class="sxs-lookup"><span data-stu-id="25499-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="25499-154">La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="25499-154">The value must be one of the following:</span></span>

- <span data-ttu-id="25499-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="25499-155">`Document` (Word)</span></span>
- <span data-ttu-id="25499-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="25499-156">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="25499-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="25499-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="25499-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="25499-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="25499-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="25499-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="25499-160">Exemple d’hôte</span><span class="sxs-lookup"><span data-stu-id="25499-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
