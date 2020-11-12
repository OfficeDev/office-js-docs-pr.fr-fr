---
title: Élément OfficeApp dans le fichier manifeste
description: L’élément OfficeApp est l’élément racine d’un manifeste de complément Office.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: c5786343173d0e130df4b786f28a8689d573b6ca
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996318"
---
# <a name="officeapp-element"></a><span data-ttu-id="3418e-103">OfficeApp, élément</span><span class="sxs-lookup"><span data-stu-id="3418e-103">OfficeApp element</span></span>

<span data-ttu-id="3418e-104">Élément racine dans le manifeste d’un complément Office.</span><span class="sxs-lookup"><span data-stu-id="3418e-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="3418e-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="3418e-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3418e-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="3418e-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="3418e-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="3418e-107">Contained in</span></span>

 <span data-ttu-id="3418e-108">_none_</span><span class="sxs-lookup"><span data-stu-id="3418e-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="3418e-109">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="3418e-109">Must contain</span></span>

|<span data-ttu-id="3418e-110">Élément</span><span class="sxs-lookup"><span data-stu-id="3418e-110">Element</span></span>|<span data-ttu-id="3418e-111">Contenu</span><span class="sxs-lookup"><span data-stu-id="3418e-111">Content</span></span>|<span data-ttu-id="3418e-112">Courrier</span><span class="sxs-lookup"><span data-stu-id="3418e-112">Mail</span></span>|<span data-ttu-id="3418e-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="3418e-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="3418e-114">Id</span><span class="sxs-lookup"><span data-stu-id="3418e-114">Id</span></span>](id.md)|<span data-ttu-id="3418e-115">x</span><span class="sxs-lookup"><span data-stu-id="3418e-115">x</span></span>|<span data-ttu-id="3418e-116">x</span><span class="sxs-lookup"><span data-stu-id="3418e-116">x</span></span>|<span data-ttu-id="3418e-117">x</span><span class="sxs-lookup"><span data-stu-id="3418e-117">x</span></span>|
|[<span data-ttu-id="3418e-118">Version</span><span class="sxs-lookup"><span data-stu-id="3418e-118">Version</span></span>](version.md)|<span data-ttu-id="3418e-119">x</span><span class="sxs-lookup"><span data-stu-id="3418e-119">x</span></span>|<span data-ttu-id="3418e-120">x</span><span class="sxs-lookup"><span data-stu-id="3418e-120">x</span></span>|<span data-ttu-id="3418e-121">x</span><span class="sxs-lookup"><span data-stu-id="3418e-121">x</span></span>|
|[<span data-ttu-id="3418e-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="3418e-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="3418e-123">x</span><span class="sxs-lookup"><span data-stu-id="3418e-123">x</span></span>|<span data-ttu-id="3418e-124">x</span><span class="sxs-lookup"><span data-stu-id="3418e-124">x</span></span>|<span data-ttu-id="3418e-125">x</span><span class="sxs-lookup"><span data-stu-id="3418e-125">x</span></span>|
|[<span data-ttu-id="3418e-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="3418e-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="3418e-127">x</span><span class="sxs-lookup"><span data-stu-id="3418e-127">x</span></span>|<span data-ttu-id="3418e-128">x</span><span class="sxs-lookup"><span data-stu-id="3418e-128">x</span></span>|<span data-ttu-id="3418e-129">x</span><span class="sxs-lookup"><span data-stu-id="3418e-129">x</span></span>|
|[<span data-ttu-id="3418e-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="3418e-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="3418e-131">x</span><span class="sxs-lookup"><span data-stu-id="3418e-131">x</span></span>||<span data-ttu-id="3418e-132">x</span><span class="sxs-lookup"><span data-stu-id="3418e-132">x</span></span>|
|[<span data-ttu-id="3418e-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="3418e-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="3418e-134">x</span><span class="sxs-lookup"><span data-stu-id="3418e-134">x</span></span>|<span data-ttu-id="3418e-135">x</span><span class="sxs-lookup"><span data-stu-id="3418e-135">x</span></span>|<span data-ttu-id="3418e-136">x</span><span class="sxs-lookup"><span data-stu-id="3418e-136">x</span></span>|
|[<span data-ttu-id="3418e-137">Description</span><span class="sxs-lookup"><span data-stu-id="3418e-137">Description</span></span>](description.md)|<span data-ttu-id="3418e-138">x</span><span class="sxs-lookup"><span data-stu-id="3418e-138">x</span></span>|<span data-ttu-id="3418e-139">x</span><span class="sxs-lookup"><span data-stu-id="3418e-139">x</span></span>|<span data-ttu-id="3418e-140">x</span><span class="sxs-lookup"><span data-stu-id="3418e-140">x</span></span>|
|[<span data-ttu-id="3418e-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="3418e-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="3418e-142">x</span><span class="sxs-lookup"><span data-stu-id="3418e-142">x</span></span>||
|[<span data-ttu-id="3418e-143">Permissions</span><span class="sxs-lookup"><span data-stu-id="3418e-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="3418e-144">x</span><span class="sxs-lookup"><span data-stu-id="3418e-144">x</span></span>||<span data-ttu-id="3418e-145">x</span><span class="sxs-lookup"><span data-stu-id="3418e-145">x</span></span>|
|[<span data-ttu-id="3418e-146">Règle</span><span class="sxs-lookup"><span data-stu-id="3418e-146">Rule</span></span>](rule.md)||<span data-ttu-id="3418e-147">x</span><span class="sxs-lookup"><span data-stu-id="3418e-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="3418e-148">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="3418e-148">Can contain</span></span>

|<span data-ttu-id="3418e-149">Élément</span><span class="sxs-lookup"><span data-stu-id="3418e-149">Element</span></span>|<span data-ttu-id="3418e-150">Contenu</span><span class="sxs-lookup"><span data-stu-id="3418e-150">Content</span></span>|<span data-ttu-id="3418e-151">Courrier</span><span class="sxs-lookup"><span data-stu-id="3418e-151">Mail</span></span>|<span data-ttu-id="3418e-152">TaskPane</span><span class="sxs-lookup"><span data-stu-id="3418e-152">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="3418e-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="3418e-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="3418e-154">x</span><span class="sxs-lookup"><span data-stu-id="3418e-154">x</span></span>|<span data-ttu-id="3418e-155">x</span><span class="sxs-lookup"><span data-stu-id="3418e-155">x</span></span>|<span data-ttu-id="3418e-156">x</span><span class="sxs-lookup"><span data-stu-id="3418e-156">x</span></span>|
|[<span data-ttu-id="3418e-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="3418e-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="3418e-158">x</span><span class="sxs-lookup"><span data-stu-id="3418e-158">x</span></span>|<span data-ttu-id="3418e-159">x</span><span class="sxs-lookup"><span data-stu-id="3418e-159">x</span></span>|<span data-ttu-id="3418e-160">x</span><span class="sxs-lookup"><span data-stu-id="3418e-160">x</span></span>|
|[<span data-ttu-id="3418e-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="3418e-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="3418e-162">x</span><span class="sxs-lookup"><span data-stu-id="3418e-162">x</span></span>|<span data-ttu-id="3418e-163">x</span><span class="sxs-lookup"><span data-stu-id="3418e-163">x</span></span>|<span data-ttu-id="3418e-164">x</span><span class="sxs-lookup"><span data-stu-id="3418e-164">x</span></span>|
|[<span data-ttu-id="3418e-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="3418e-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="3418e-166">x</span><span class="sxs-lookup"><span data-stu-id="3418e-166">x</span></span>|<span data-ttu-id="3418e-167">x</span><span class="sxs-lookup"><span data-stu-id="3418e-167">x</span></span>|<span data-ttu-id="3418e-168">x</span><span class="sxs-lookup"><span data-stu-id="3418e-168">x</span></span>|
|[<span data-ttu-id="3418e-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="3418e-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="3418e-170">x</span><span class="sxs-lookup"><span data-stu-id="3418e-170">x</span></span>|<span data-ttu-id="3418e-171">x</span><span class="sxs-lookup"><span data-stu-id="3418e-171">x</span></span>|<span data-ttu-id="3418e-172">x</span><span class="sxs-lookup"><span data-stu-id="3418e-172">x</span></span>|
|[<span data-ttu-id="3418e-173">Hôtes</span><span class="sxs-lookup"><span data-stu-id="3418e-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="3418e-174">x</span><span class="sxs-lookup"><span data-stu-id="3418e-174">x</span></span>|<span data-ttu-id="3418e-175">x</span><span class="sxs-lookup"><span data-stu-id="3418e-175">x</span></span>|<span data-ttu-id="3418e-176">x</span><span class="sxs-lookup"><span data-stu-id="3418e-176">x</span></span>|
|[<span data-ttu-id="3418e-177">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3418e-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="3418e-178">x</span><span class="sxs-lookup"><span data-stu-id="3418e-178">x</span></span>|<span data-ttu-id="3418e-179">x</span><span class="sxs-lookup"><span data-stu-id="3418e-179">x</span></span>|<span data-ttu-id="3418e-180">x</span><span class="sxs-lookup"><span data-stu-id="3418e-180">x</span></span>|
|[<span data-ttu-id="3418e-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="3418e-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="3418e-182">x</span><span class="sxs-lookup"><span data-stu-id="3418e-182">x</span></span>|||
|[<span data-ttu-id="3418e-183">Permissions</span><span class="sxs-lookup"><span data-stu-id="3418e-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="3418e-184">x</span><span class="sxs-lookup"><span data-stu-id="3418e-184">x</span></span>||
|[<span data-ttu-id="3418e-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="3418e-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="3418e-186">x</span><span class="sxs-lookup"><span data-stu-id="3418e-186">x</span></span>||
|[<span data-ttu-id="3418e-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="3418e-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="3418e-188">x</span><span class="sxs-lookup"><span data-stu-id="3418e-188">x</span></span>|
|[<span data-ttu-id="3418e-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="3418e-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="3418e-190">x</span><span class="sxs-lookup"><span data-stu-id="3418e-190">x</span></span>|<span data-ttu-id="3418e-191">x</span><span class="sxs-lookup"><span data-stu-id="3418e-191">x</span></span>|<span data-ttu-id="3418e-192">x</span><span class="sxs-lookup"><span data-stu-id="3418e-192">x</span></span>|
|[<span data-ttu-id="3418e-193">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="3418e-193">ExtendedOverrides</span></span>](extendedoverrides.md)|||<span data-ttu-id="3418e-194">x</span><span class="sxs-lookup"><span data-stu-id="3418e-194">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="3418e-195">Attributs</span><span class="sxs-lookup"><span data-stu-id="3418e-195">Attributes</span></span>

|<span data-ttu-id="3418e-196">Attribut</span><span class="sxs-lookup"><span data-stu-id="3418e-196">Attribute</span></span>|<span data-ttu-id="3418e-197">Description</span><span class="sxs-lookup"><span data-stu-id="3418e-197">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="3418e-198">xmlns</span><span class="sxs-lookup"><span data-stu-id="3418e-198">xmlns</span></span>|<span data-ttu-id="3418e-p101">Définit la version de schéma et l’espace de noms du manifeste de complément Office. Cet attribut doit toujours être défini sur `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="3418e-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="3418e-201">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="3418e-201">xmlns:xsi</span></span>|<span data-ttu-id="3418e-p102">Définit l’instance XMLSchema. Cet attribut doit toujours être défini sur `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="3418e-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="3418e-204">xsi:type</span><span class="sxs-lookup"><span data-stu-id="3418e-204">xsi:type</span></span>|<span data-ttu-id="3418e-p103">Définit le type de complément Office. Cet attribut doit être défini sur l’une des options suivantes : `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`.</span><span class="sxs-lookup"><span data-stu-id="3418e-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
