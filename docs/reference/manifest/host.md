---
title: Élément Host dans le fichier manifeste
description: ''
ms.date: 07/01/2019
localization_priority: Normal
ms.openlocfilehash: e7b557034f70b03ed57598b7ffb9f43878db7392
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454894"
---
# <a name="host-element"></a><span data-ttu-id="5cc81-102">Élément Host</span><span class="sxs-lookup"><span data-stu-id="5cc81-102">Host element</span></span>

<span data-ttu-id="5cc81-103">Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.</span><span class="sxs-lookup"><span data-stu-id="5cc81-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="5cc81-104">La syntaxe des éléments **Host** varie selon que l’élément est défini dans le [manifeste de base](#basic-manifest) ou le nœud [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="5cc81-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="5cc81-105">Toutefois, la fonctionnalité est identique.</span><span class="sxs-lookup"><span data-stu-id="5cc81-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="5cc81-106">Manifeste de base</span><span class="sxs-lookup"><span data-stu-id="5cc81-106">Basic manifest</span></span>

<span data-ttu-id="5cc81-107">Lorsqu’il est défini dans le manifeste base (sous [OfficeApp](officeapp.md)), le type d’hôte est déterminé par l’attribut `Name`.</span><span class="sxs-lookup"><span data-stu-id="5cc81-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="5cc81-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="5cc81-108">Attributes</span></span>

| <span data-ttu-id="5cc81-109">Attribut</span><span class="sxs-lookup"><span data-stu-id="5cc81-109">Attribute</span></span>     | <span data-ttu-id="5cc81-110">Type</span><span class="sxs-lookup"><span data-stu-id="5cc81-110">Type</span></span>   | <span data-ttu-id="5cc81-111">Requis</span><span class="sxs-lookup"><span data-stu-id="5cc81-111">Required</span></span> | <span data-ttu-id="5cc81-112">Description</span><span class="sxs-lookup"><span data-stu-id="5cc81-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="5cc81-113">Name</span><span class="sxs-lookup"><span data-stu-id="5cc81-113">Name</span></span>](#name) | <span data-ttu-id="5cc81-114">string</span><span class="sxs-lookup"><span data-stu-id="5cc81-114">string</span></span> | <span data-ttu-id="5cc81-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="5cc81-115">required</span></span> | <span data-ttu-id="5cc81-116">Nom du type d’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="5cc81-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="5cc81-117">Nom</span><span class="sxs-lookup"><span data-stu-id="5cc81-117">Name</span></span>

<span data-ttu-id="5cc81-118">Spécifie le type d’hôte ciblé par ce complément.</span><span class="sxs-lookup"><span data-stu-id="5cc81-118">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="5cc81-119">La valeur doit être l’une des valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="5cc81-119">The value must be one of the following.</span></span>

- <span data-ttu-id="5cc81-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="5cc81-120">`Document` (Word)</span></span>
- <span data-ttu-id="5cc81-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="5cc81-121">`Database` (Access)</span></span>
- <span data-ttu-id="5cc81-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="5cc81-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="5cc81-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="5cc81-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="5cc81-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="5cc81-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="5cc81-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="5cc81-125">`Project` (Project)</span></span>
- <span data-ttu-id="5cc81-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="5cc81-126">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5cc81-127">Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint.</span><span class="sxs-lookup"><span data-stu-id="5cc81-127">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="5cc81-128">Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.</span><span class="sxs-lookup"><span data-stu-id="5cc81-128">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="5cc81-129">Exemple</span><span class="sxs-lookup"><span data-stu-id="5cc81-129">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="5cc81-130">Nœud VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="5cc81-130">VersionOverrides node</span></span>

<span data-ttu-id="5cc81-131">Lorsqu’il est défini dans [VersionOverrides](versionoverrides.md), le type d’hôte est déterminé par l’attribut `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="5cc81-131">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="5cc81-132">Attributs</span><span class="sxs-lookup"><span data-stu-id="5cc81-132">Attributes</span></span>

|  <span data-ttu-id="5cc81-133">Attribut</span><span class="sxs-lookup"><span data-stu-id="5cc81-133">Attribute</span></span>  |  <span data-ttu-id="5cc81-134">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="5cc81-134">Required</span></span>  |  <span data-ttu-id="5cc81-135">Description</span><span class="sxs-lookup"><span data-stu-id="5cc81-135">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5cc81-136">xsi:type</span><span class="sxs-lookup"><span data-stu-id="5cc81-136">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="5cc81-137">Oui</span><span class="sxs-lookup"><span data-stu-id="5cc81-137">Yes</span></span>  | <span data-ttu-id="5cc81-138">Décrit l’hôte d’Office dans lequel ces paramètres s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="5cc81-138">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="5cc81-139">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="5cc81-139">Child elements</span></span>

|  <span data-ttu-id="5cc81-140">Élément</span><span class="sxs-lookup"><span data-stu-id="5cc81-140">Element</span></span> |  <span data-ttu-id="5cc81-141">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="5cc81-141">Required</span></span>  |  <span data-ttu-id="5cc81-142">Description</span><span class="sxs-lookup"><span data-stu-id="5cc81-142">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5cc81-143">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="5cc81-143">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="5cc81-144">Oui</span><span class="sxs-lookup"><span data-stu-id="5cc81-144">Yes</span></span>   |  <span data-ttu-id="5cc81-145">Définit les paramètres pour le facteur de forme pour bureau.</span><span class="sxs-lookup"><span data-stu-id="5cc81-145">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="5cc81-146">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="5cc81-146">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="5cc81-147">Non</span><span class="sxs-lookup"><span data-stu-id="5cc81-147">No</span></span>   |  <span data-ttu-id="5cc81-148">Définit les paramètres pour le facteur de forme pour environnement mobile.</span><span class="sxs-lookup"><span data-stu-id="5cc81-148">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="5cc81-149">**Remarque:** Cet élément est pris en charge uniquement dans Outlook sur iOS.</span><span class="sxs-lookup"><span data-stu-id="5cc81-149">**Note:** This element is only supported in Outlook on iOS.</span></span> |
|  [<span data-ttu-id="5cc81-150">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="5cc81-150">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="5cc81-151">Non</span><span class="sxs-lookup"><span data-stu-id="5cc81-151">No</span></span>   |  <span data-ttu-id="5cc81-152">Définit les paramètres de tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="5cc81-152">Defines the settings for all form factors.</span></span> <span data-ttu-id="5cc81-153">Utilisé uniquement par des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="5cc81-153">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="5cc81-154">xsi:type</span><span class="sxs-lookup"><span data-stu-id="5cc81-154">xsi:type</span></span>

<span data-ttu-id="5cc81-155">Contrôle à quel hôte Office (Word, Excel, PowerPoint, Outlook, OneNote) s’applique également les paramètres contenus.</span><span class="sxs-lookup"><span data-stu-id="5cc81-155">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="5cc81-156">La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="5cc81-156">The value must be one of the following:</span></span>

- <span data-ttu-id="5cc81-157">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="5cc81-157">`Document` (Word)</span></span>
- <span data-ttu-id="5cc81-158">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="5cc81-158">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="5cc81-159">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="5cc81-159">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="5cc81-160">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="5cc81-160">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="5cc81-161">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="5cc81-161">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="5cc81-162">Exemple d’hôte</span><span class="sxs-lookup"><span data-stu-id="5cc81-162">Host example</span></span> 

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
