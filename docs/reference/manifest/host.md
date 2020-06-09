---
title: Élément Host dans le fichier manifeste
description: Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5db9df97c4ba558d54756b983a26cb7b71e049d5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611812"
---
# <a name="host-element"></a><span data-ttu-id="ac77b-103">Élément Host</span><span class="sxs-lookup"><span data-stu-id="ac77b-103">Host element</span></span>

<span data-ttu-id="ac77b-104">Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.</span><span class="sxs-lookup"><span data-stu-id="ac77b-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ac77b-105">La syntaxe des éléments **Host** varie selon que l’élément est défini dans le [manifeste de base](#basic-manifest) ou le nœud [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="ac77b-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="ac77b-106">Toutefois, la fonctionnalité est identique.</span><span class="sxs-lookup"><span data-stu-id="ac77b-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="ac77b-107">Manifeste de base</span><span class="sxs-lookup"><span data-stu-id="ac77b-107">Basic manifest</span></span>

<span data-ttu-id="ac77b-108">Lorsqu’il est défini dans le manifeste base (sous [OfficeApp](officeapp.md)), le type d’hôte est déterminé par l’attribut `Name`.</span><span class="sxs-lookup"><span data-stu-id="ac77b-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="ac77b-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="ac77b-109">Attributes</span></span>

| <span data-ttu-id="ac77b-110">Attribut</span><span class="sxs-lookup"><span data-stu-id="ac77b-110">Attribute</span></span>     | <span data-ttu-id="ac77b-111">Type</span><span class="sxs-lookup"><span data-stu-id="ac77b-111">Type</span></span>   | <span data-ttu-id="ac77b-112">Requis</span><span class="sxs-lookup"><span data-stu-id="ac77b-112">Required</span></span> | <span data-ttu-id="ac77b-113">Description</span><span class="sxs-lookup"><span data-stu-id="ac77b-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="ac77b-114">Name</span><span class="sxs-lookup"><span data-stu-id="ac77b-114">Name</span></span>](#name) | <span data-ttu-id="ac77b-115">string</span><span class="sxs-lookup"><span data-stu-id="ac77b-115">string</span></span> | <span data-ttu-id="ac77b-116">obligatoire</span><span class="sxs-lookup"><span data-stu-id="ac77b-116">required</span></span> | <span data-ttu-id="ac77b-117">Nom du type d’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="ac77b-117">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="ac77b-118">Nom</span><span class="sxs-lookup"><span data-stu-id="ac77b-118">Name</span></span>

<span data-ttu-id="ac77b-119">Spécifie le type d’hôte ciblé par ce complément.</span><span class="sxs-lookup"><span data-stu-id="ac77b-119">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="ac77b-120">La valeur doit être l’une des valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="ac77b-120">The value must be one of the following.</span></span>

- <span data-ttu-id="ac77b-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="ac77b-121">`Document` (Word)</span></span>
- <span data-ttu-id="ac77b-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="ac77b-122">`Database` (Access)</span></span>
- <span data-ttu-id="ac77b-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="ac77b-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="ac77b-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="ac77b-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="ac77b-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="ac77b-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="ac77b-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="ac77b-126">`Project` (Project)</span></span>
- <span data-ttu-id="ac77b-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="ac77b-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ac77b-128">Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint.</span><span class="sxs-lookup"><span data-stu-id="ac77b-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="ac77b-129">Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.</span><span class="sxs-lookup"><span data-stu-id="ac77b-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="ac77b-130">Exemple</span><span class="sxs-lookup"><span data-stu-id="ac77b-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="ac77b-131">Nœud VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="ac77b-131">VersionOverrides node</span></span>

<span data-ttu-id="ac77b-132">Lorsqu’il est défini dans [VersionOverrides](versionoverrides.md), le type d’hôte est déterminé par l’attribut `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="ac77b-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="ac77b-133">Attributs</span><span class="sxs-lookup"><span data-stu-id="ac77b-133">Attributes</span></span>

|  <span data-ttu-id="ac77b-134">Attribut</span><span class="sxs-lookup"><span data-stu-id="ac77b-134">Attribute</span></span>  |  <span data-ttu-id="ac77b-135">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="ac77b-135">Required</span></span>  |  <span data-ttu-id="ac77b-136">Description</span><span class="sxs-lookup"><span data-stu-id="ac77b-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ac77b-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ac77b-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="ac77b-138">Oui</span><span class="sxs-lookup"><span data-stu-id="ac77b-138">Yes</span></span>  | <span data-ttu-id="ac77b-139">Décrit l’hôte d’Office dans lequel ces paramètres s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="ac77b-139">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="ac77b-140">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="ac77b-140">Child elements</span></span>

|  <span data-ttu-id="ac77b-141">Élément</span><span class="sxs-lookup"><span data-stu-id="ac77b-141">Element</span></span> |  <span data-ttu-id="ac77b-142">Requis</span><span class="sxs-lookup"><span data-stu-id="ac77b-142">Required</span></span>  |  <span data-ttu-id="ac77b-143">Description</span><span class="sxs-lookup"><span data-stu-id="ac77b-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ac77b-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="ac77b-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="ac77b-145">Oui</span><span class="sxs-lookup"><span data-stu-id="ac77b-145">Yes</span></span>   |  <span data-ttu-id="ac77b-146">Définit les paramètres pour le facteur de forme pour bureau.</span><span class="sxs-lookup"><span data-stu-id="ac77b-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="ac77b-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="ac77b-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="ac77b-148">Non</span><span class="sxs-lookup"><span data-stu-id="ac77b-148">No</span></span>   |  <span data-ttu-id="ac77b-149">Définit les paramètres pour le facteur de forme pour environnement mobile.</span><span class="sxs-lookup"><span data-stu-id="ac77b-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="ac77b-150">**Remarque :** Cet élément est pris en charge uniquement dans Outlook sur iOS et Android.</span><span class="sxs-lookup"><span data-stu-id="ac77b-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="ac77b-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="ac77b-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="ac77b-152">Non</span><span class="sxs-lookup"><span data-stu-id="ac77b-152">No</span></span>   |  <span data-ttu-id="ac77b-153">Définit les paramètres de tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="ac77b-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="ac77b-154">Utilisé uniquement par des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="ac77b-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="ac77b-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ac77b-155">xsi:type</span></span>

<span data-ttu-id="ac77b-156">Contrôle à quel hôte Office (Word, Excel, PowerPoint, Outlook, OneNote) s’applique également les paramètres contenus.</span><span class="sxs-lookup"><span data-stu-id="ac77b-156">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="ac77b-157">La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="ac77b-157">The value must be one of the following:</span></span>

- <span data-ttu-id="ac77b-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="ac77b-158">`Document` (Word)</span></span>
- <span data-ttu-id="ac77b-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="ac77b-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="ac77b-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="ac77b-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="ac77b-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="ac77b-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="ac77b-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="ac77b-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="ac77b-163">Exemple d’hôte</span><span class="sxs-lookup"><span data-stu-id="ac77b-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
