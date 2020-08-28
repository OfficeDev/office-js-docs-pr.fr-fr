---
title: Élément Host dans le fichier manifeste
description: Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5b6c6e6b5471b4117c28cf92e11eb0a99b512a97
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292285"
---
# <a name="host-element"></a><span data-ttu-id="9b784-103">Élément Host</span><span class="sxs-lookup"><span data-stu-id="9b784-103">Host element</span></span>

<span data-ttu-id="9b784-104">Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.</span><span class="sxs-lookup"><span data-stu-id="9b784-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9b784-105">La syntaxe des éléments **Host** varie selon que l’élément est défini dans le [manifeste de base](#basic-manifest) ou le nœud [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="9b784-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="9b784-106">Toutefois, la fonctionnalité est identique.</span><span class="sxs-lookup"><span data-stu-id="9b784-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="9b784-107">Manifeste de base</span><span class="sxs-lookup"><span data-stu-id="9b784-107">Basic manifest</span></span>

<span data-ttu-id="9b784-108">Lorsqu’il est défini dans le manifeste base (sous [OfficeApp](officeapp.md)), le type d’hôte est déterminé par l’attribut `Name`.</span><span class="sxs-lookup"><span data-stu-id="9b784-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="9b784-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="9b784-109">Attributes</span></span>

| <span data-ttu-id="9b784-110">Attribut</span><span class="sxs-lookup"><span data-stu-id="9b784-110">Attribute</span></span>     | <span data-ttu-id="9b784-111">Type</span><span class="sxs-lookup"><span data-stu-id="9b784-111">Type</span></span>   | <span data-ttu-id="9b784-112">Requis</span><span class="sxs-lookup"><span data-stu-id="9b784-112">Required</span></span> | <span data-ttu-id="9b784-113">Description</span><span class="sxs-lookup"><span data-stu-id="9b784-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="9b784-114">Name</span><span class="sxs-lookup"><span data-stu-id="9b784-114">Name</span></span>](#name) | <span data-ttu-id="9b784-115">string</span><span class="sxs-lookup"><span data-stu-id="9b784-115">string</span></span> | <span data-ttu-id="9b784-116">obligatoire</span><span class="sxs-lookup"><span data-stu-id="9b784-116">required</span></span> | <span data-ttu-id="9b784-117">Nom du type d’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="9b784-117">The name of the type of Office client application.</span></span> |

### <a name="name"></a><span data-ttu-id="9b784-118">Nom</span><span class="sxs-lookup"><span data-stu-id="9b784-118">Name</span></span>

<span data-ttu-id="9b784-119">Spécifie le type d’hôte ciblé par ce complément.</span><span class="sxs-lookup"><span data-stu-id="9b784-119">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="9b784-120">La valeur doit être l’une des valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="9b784-120">The value must be one of the following.</span></span>

- <span data-ttu-id="9b784-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="9b784-121">`Document` (Word)</span></span>
- <span data-ttu-id="9b784-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="9b784-122">`Database` (Access)</span></span>
- <span data-ttu-id="9b784-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="9b784-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="9b784-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="9b784-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="9b784-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="9b784-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="9b784-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="9b784-126">`Project` (Project)</span></span>
- <span data-ttu-id="9b784-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="9b784-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9b784-128">Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint.</span><span class="sxs-lookup"><span data-stu-id="9b784-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="9b784-129">Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.</span><span class="sxs-lookup"><span data-stu-id="9b784-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="9b784-130">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b784-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="9b784-131">Nœud VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="9b784-131">VersionOverrides node</span></span>

<span data-ttu-id="9b784-132">Lorsqu’il est défini dans [VersionOverrides](versionoverrides.md), le type d’hôte est déterminé par l’attribut `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="9b784-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="9b784-133">Attributs</span><span class="sxs-lookup"><span data-stu-id="9b784-133">Attributes</span></span>

|  <span data-ttu-id="9b784-134">Attribut</span><span class="sxs-lookup"><span data-stu-id="9b784-134">Attribute</span></span>  |  <span data-ttu-id="9b784-135">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="9b784-135">Required</span></span>  |  <span data-ttu-id="9b784-136">Description</span><span class="sxs-lookup"><span data-stu-id="9b784-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9b784-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9b784-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="9b784-138">Oui</span><span class="sxs-lookup"><span data-stu-id="9b784-138">Yes</span></span>  | <span data-ttu-id="9b784-139">Décrit l’application Office à laquelle ces paramètres s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="9b784-139">Describes the Office application where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="9b784-140">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="9b784-140">Child elements</span></span>

|  <span data-ttu-id="9b784-141">Élément</span><span class="sxs-lookup"><span data-stu-id="9b784-141">Element</span></span> |  <span data-ttu-id="9b784-142">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="9b784-142">Required</span></span>  |  <span data-ttu-id="9b784-143">Description</span><span class="sxs-lookup"><span data-stu-id="9b784-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9b784-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="9b784-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="9b784-145">Oui</span><span class="sxs-lookup"><span data-stu-id="9b784-145">Yes</span></span>   |  <span data-ttu-id="9b784-146">Définit les paramètres pour le facteur de forme pour bureau.</span><span class="sxs-lookup"><span data-stu-id="9b784-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="9b784-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="9b784-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="9b784-148">Non</span><span class="sxs-lookup"><span data-stu-id="9b784-148">No</span></span>   |  <span data-ttu-id="9b784-149">Définit les paramètres pour le facteur de forme pour environnement mobile.</span><span class="sxs-lookup"><span data-stu-id="9b784-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="9b784-150">**Remarque :** Cet élément est pris en charge uniquement dans Outlook sur iOS et Android.</span><span class="sxs-lookup"><span data-stu-id="9b784-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="9b784-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="9b784-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="9b784-152">Non</span><span class="sxs-lookup"><span data-stu-id="9b784-152">No</span></span>   |  <span data-ttu-id="9b784-153">Définit les paramètres de tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="9b784-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="9b784-154">Utilisé uniquement par des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="9b784-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="9b784-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9b784-155">xsi:type</span></span>

<span data-ttu-id="9b784-156">Détermine l’application Office (Word, Excel, PowerPoint, Outlook, OneNote) à laquelle les paramètres contenus s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="9b784-156">Controls which Office application (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="9b784-157">La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="9b784-157">The value must be one of the following:</span></span>

- <span data-ttu-id="9b784-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="9b784-158">`Document` (Word)</span></span>
- <span data-ttu-id="9b784-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="9b784-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="9b784-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="9b784-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="9b784-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="9b784-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="9b784-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="9b784-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="9b784-163">Exemple d’hôte</span><span class="sxs-lookup"><span data-stu-id="9b784-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
