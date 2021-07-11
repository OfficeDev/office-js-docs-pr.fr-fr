---
title: Élément Host dans le fichier manifeste
description: Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 45d4ed42946038699be235ff3912c071a92ff226
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348327"
---
# <a name="host-element"></a><span data-ttu-id="fe3e2-103">Élément Host</span><span class="sxs-lookup"><span data-stu-id="fe3e2-103">Host element</span></span>

<span data-ttu-id="fe3e2-104">Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fe3e2-105">La syntaxe des éléments **Host** varie selon que l’élément est défini dans le [manifeste de base](#basic-manifest) ou le nœud [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="fe3e2-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="fe3e2-106">Toutefois, la fonctionnalité est identique.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="fe3e2-107">Manifeste de base</span><span class="sxs-lookup"><span data-stu-id="fe3e2-107">Basic manifest</span></span>

<span data-ttu-id="fe3e2-108">Lorsqu’il est défini dans le manifeste base (sous [OfficeApp](officeapp.md)), le type d’hôte est déterminé par l’attribut `Name`.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="fe3e2-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="fe3e2-109">Attributes</span></span>

| <span data-ttu-id="fe3e2-110">Attribut</span><span class="sxs-lookup"><span data-stu-id="fe3e2-110">Attribute</span></span>     | <span data-ttu-id="fe3e2-111">Type</span><span class="sxs-lookup"><span data-stu-id="fe3e2-111">Type</span></span>   | <span data-ttu-id="fe3e2-112">Requis</span><span class="sxs-lookup"><span data-stu-id="fe3e2-112">Required</span></span> | <span data-ttu-id="fe3e2-113">Description</span><span class="sxs-lookup"><span data-stu-id="fe3e2-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="fe3e2-114">Name</span><span class="sxs-lookup"><span data-stu-id="fe3e2-114">Name</span></span>](#name) | <span data-ttu-id="fe3e2-115">string</span><span class="sxs-lookup"><span data-stu-id="fe3e2-115">string</span></span> | <span data-ttu-id="fe3e2-116">obligatoire</span><span class="sxs-lookup"><span data-stu-id="fe3e2-116">required</span></span> | <span data-ttu-id="fe3e2-117">Nom du type d’application Office client.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-117">The name of the type of Office client application.</span></span> |

### <a name="name"></a><span data-ttu-id="fe3e2-118">Nom</span><span class="sxs-lookup"><span data-stu-id="fe3e2-118">Name</span></span>

<span data-ttu-id="fe3e2-p102">Spécifie le type d’hôte ciblé par ce complément. La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="fe3e2-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="fe3e2-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-121">`Document` (Word)</span></span>
- <span data-ttu-id="fe3e2-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-122">`Database` (Access)</span></span>
- <span data-ttu-id="fe3e2-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="fe3e2-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="fe3e2-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="fe3e2-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-126">`Project` (Project)</span></span>
- <span data-ttu-id="fe3e2-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fe3e2-128">Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="fe3e2-129">Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="fe3e2-130">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe3e2-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="fe3e2-131">Nœud VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="fe3e2-131">VersionOverrides node</span></span>

<span data-ttu-id="fe3e2-132">Lorsqu’il est défini dans [VersionOverrides](versionoverrides.md), le type d’hôte est déterminé par l’attribut `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="fe3e2-133">Attributs</span><span class="sxs-lookup"><span data-stu-id="fe3e2-133">Attributes</span></span>

|  <span data-ttu-id="fe3e2-134">Attribut</span><span class="sxs-lookup"><span data-stu-id="fe3e2-134">Attribute</span></span>  |  <span data-ttu-id="fe3e2-135">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="fe3e2-135">Required</span></span>  |  <span data-ttu-id="fe3e2-136">Description</span><span class="sxs-lookup"><span data-stu-id="fe3e2-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fe3e2-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="fe3e2-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="fe3e2-138">Oui</span><span class="sxs-lookup"><span data-stu-id="fe3e2-138">Yes</span></span>  | <span data-ttu-id="fe3e2-139">Décrit l’application Office application dans laquelle ces paramètres s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-139">Describes the Office application where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="fe3e2-140">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="fe3e2-140">Child elements</span></span>

|  <span data-ttu-id="fe3e2-141">Élément</span><span class="sxs-lookup"><span data-stu-id="fe3e2-141">Element</span></span> |  <span data-ttu-id="fe3e2-142">Requis</span><span class="sxs-lookup"><span data-stu-id="fe3e2-142">Required</span></span>  |  <span data-ttu-id="fe3e2-143">Description</span><span class="sxs-lookup"><span data-stu-id="fe3e2-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fe3e2-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="fe3e2-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="fe3e2-145">Oui</span><span class="sxs-lookup"><span data-stu-id="fe3e2-145">Yes</span></span>   |  <span data-ttu-id="fe3e2-146">Définit les paramètres pour le facteur de forme pour bureau.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="fe3e2-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="fe3e2-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="fe3e2-148">Non</span><span class="sxs-lookup"><span data-stu-id="fe3e2-148">No</span></span>   |  <span data-ttu-id="fe3e2-149">Définit les paramètres pour le facteur de forme pour environnement mobile.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="fe3e2-150">**Remarque :** Cet élément est uniquement pris en charge dans Outlook sur iOS et Android.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="fe3e2-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="fe3e2-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="fe3e2-152">Non</span><span class="sxs-lookup"><span data-stu-id="fe3e2-152">No</span></span>   |  <span data-ttu-id="fe3e2-153">Définit les paramètres de tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="fe3e2-154">Utilisé uniquement par des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="fe3e2-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="fe3e2-155">xsi:type</span></span>

<span data-ttu-id="fe3e2-156">Contrôle l Office application (Word, Excel, PowerPoint, Outlook, OneNote) dans laquelle les paramètres contenus s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="fe3e2-156">Controls which Office application (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="fe3e2-157">La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="fe3e2-157">The value must be one of the following:</span></span>

- <span data-ttu-id="fe3e2-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-158">`Document` (Word)</span></span>
- <span data-ttu-id="fe3e2-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="fe3e2-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="fe3e2-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="fe3e2-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="fe3e2-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="fe3e2-163">Exemple d’hôte</span><span class="sxs-lookup"><span data-stu-id="fe3e2-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
