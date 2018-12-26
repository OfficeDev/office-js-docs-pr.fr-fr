---
title: Élément Host dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 37b772261ad82b4f899e73314a08ffd1dd03b442
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432808"
---
# <a name="host-element"></a><span data-ttu-id="84f12-102">Élément Host</span><span class="sxs-lookup"><span data-stu-id="84f12-102">Host element</span></span>

<span data-ttu-id="84f12-103">Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.</span><span class="sxs-lookup"><span data-stu-id="84f12-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="84f12-104">La syntaxe des éléments **Host** varie selon que l’élément est défini dans le [manifeste de base](#basic-manifest) ou le nœud [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="84f12-104">Important: The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="84f12-105">Toutefois, la fonctionnalité est identique.</span><span class="sxs-lookup"><span data-stu-id="84f12-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="84f12-106">Manifeste de base</span><span class="sxs-lookup"><span data-stu-id="84f12-106">Basic manifest</span></span>

<span data-ttu-id="84f12-107">Lorsqu’il est défini dans le manifeste base (sous [OfficeApp](officeapp.md)), le type d’hôte est déterminé par l’attribut `Name`.</span><span class="sxs-lookup"><span data-stu-id="84f12-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="84f12-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="84f12-108">Attributes</span></span>

| <span data-ttu-id="84f12-109">Attribut</span><span class="sxs-lookup"><span data-stu-id="84f12-109">Attribute</span></span>     | <span data-ttu-id="84f12-110">Type</span><span class="sxs-lookup"><span data-stu-id="84f12-110">Type</span></span>   | <span data-ttu-id="84f12-111">Requis</span><span class="sxs-lookup"><span data-stu-id="84f12-111">Required</span></span> | <span data-ttu-id="84f12-112">Description</span><span class="sxs-lookup"><span data-stu-id="84f12-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="84f12-113">Name</span><span class="sxs-lookup"><span data-stu-id="84f12-113">Name</span></span>](#name) | <span data-ttu-id="84f12-114">chaîne</span><span class="sxs-lookup"><span data-stu-id="84f12-114">string</span></span> | <span data-ttu-id="84f12-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="84f12-115">required</span></span> | <span data-ttu-id="84f12-116">Nom du type d’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="84f12-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="84f12-117">Nom</span><span class="sxs-lookup"><span data-stu-id="84f12-117">Name</span></span>
<span data-ttu-id="84f12-p102">Spécifie le type d’hôte ciblé par ce complément. La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="84f12-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="84f12-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="84f12-120">`Document` (Word)</span></span>
- <span data-ttu-id="84f12-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="84f12-121">`Database` (Access)</span></span>
- <span data-ttu-id="84f12-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="84f12-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="84f12-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="84f12-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="84f12-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="84f12-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="84f12-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="84f12-125">`Project` (Project)</span></span>
- <span data-ttu-id="84f12-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="84f12-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="84f12-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="84f12-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="84f12-128">Nœud VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="84f12-128">VersionOverrides node</span></span>
<span data-ttu-id="84f12-129">Lorsqu’il est défini dans [VersionOverrides](versionoverrides.md), le type d’hôte est déterminé par l’attribut `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="84f12-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="84f12-130">Attributs</span><span class="sxs-lookup"><span data-stu-id="84f12-130">Attributes</span></span>

|  <span data-ttu-id="84f12-131">Attribut</span><span class="sxs-lookup"><span data-stu-id="84f12-131">Attribute</span></span>  |  <span data-ttu-id="84f12-132">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="84f12-132">Required</span></span>  |  <span data-ttu-id="84f12-133">Description</span><span class="sxs-lookup"><span data-stu-id="84f12-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="84f12-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="84f12-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="84f12-135">Oui</span><span class="sxs-lookup"><span data-stu-id="84f12-135">Yes</span></span>  | <span data-ttu-id="84f12-136">Décrit l’hôte d’Office dans lequel ces paramètres s’appliquent.</span><span class="sxs-lookup"><span data-stu-id="84f12-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="84f12-137">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="84f12-137">Child elements</span></span>

|  <span data-ttu-id="84f12-138">Élément</span><span class="sxs-lookup"><span data-stu-id="84f12-138">Element</span></span> |  <span data-ttu-id="84f12-139">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="84f12-139">Required</span></span>  |  <span data-ttu-id="84f12-140">Description</span><span class="sxs-lookup"><span data-stu-id="84f12-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="84f12-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="84f12-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="84f12-142">Oui</span><span class="sxs-lookup"><span data-stu-id="84f12-142">Yes</span></span>   |  <span data-ttu-id="84f12-143">Définit les paramètres pour le facteur de forme pour bureau.</span><span class="sxs-lookup"><span data-stu-id="84f12-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="84f12-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="84f12-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="84f12-145">Non</span><span class="sxs-lookup"><span data-stu-id="84f12-145">No</span></span>   |  <span data-ttu-id="84f12-p103">Définit les paramètres pour le facteur de forme pour environnement mobile. **Remarque :** cet élément est uniquement pris en charge dans Outlook pour iOS.</span><span class="sxs-lookup"><span data-stu-id="84f12-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="84f12-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="84f12-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="84f12-149">Non</span><span class="sxs-lookup"><span data-stu-id="84f12-149">No</span></span>   |  <span data-ttu-id="84f12-150">Définit les paramètres de tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="84f12-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="84f12-151">Utilisé uniquement par des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="84f12-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="84f12-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="84f12-152">xsi:type</span></span>

<span data-ttu-id="84f12-153">Contrôle à quel hôte Office (Word, Excel, PowerPoint, Outlook, OneNote) s’applique également les paramètres contenus.</span><span class="sxs-lookup"><span data-stu-id="84f12-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="84f12-154">La valeur doit être l’une des suivantes :</span><span class="sxs-lookup"><span data-stu-id="84f12-154">The value must be one of the following:</span></span>

- <span data-ttu-id="84f12-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="84f12-155">`Document` (Word)</span></span>
- <span data-ttu-id="84f12-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="84f12-156">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="84f12-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="84f12-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="84f12-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="84f12-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="84f12-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="84f12-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="84f12-160">Exemple d’hôte</span><span class="sxs-lookup"><span data-stu-id="84f12-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
