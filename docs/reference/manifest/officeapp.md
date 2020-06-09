---
title: Élément OfficeApp dans le fichier manifeste
description: L’élément OfficeApp est l’élément racine d’un manifeste de complément Office.
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: b6f3102a97794a19366b06734789e01fc4bc4f9d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611525"
---
# <a name="officeapp-element"></a><span data-ttu-id="c4139-103">OfficeApp, élément</span><span class="sxs-lookup"><span data-stu-id="c4139-103">OfficeApp element</span></span>

<span data-ttu-id="c4139-104">Élément racine dans le manifeste d’un complément Office.</span><span class="sxs-lookup"><span data-stu-id="c4139-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="c4139-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="c4139-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c4139-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="c4139-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="c4139-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="c4139-107">Contained in</span></span>

 <span data-ttu-id="c4139-108">_none_</span><span class="sxs-lookup"><span data-stu-id="c4139-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="c4139-109">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="c4139-109">Must contain</span></span>

|<span data-ttu-id="c4139-110">**Élément**</span><span class="sxs-lookup"><span data-stu-id="c4139-110">**Element**</span></span>|<span data-ttu-id="c4139-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="c4139-111">**Content**</span></span>|<span data-ttu-id="c4139-112">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="c4139-112">**Mail**</span></span>|<span data-ttu-id="c4139-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="c4139-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="c4139-114">Id</span><span class="sxs-lookup"><span data-stu-id="c4139-114">Id</span></span>](id.md)|<span data-ttu-id="c4139-115">x</span><span class="sxs-lookup"><span data-stu-id="c4139-115">x</span></span>|<span data-ttu-id="c4139-116">x</span><span class="sxs-lookup"><span data-stu-id="c4139-116">x</span></span>|<span data-ttu-id="c4139-117">x</span><span class="sxs-lookup"><span data-stu-id="c4139-117">x</span></span>|
|[<span data-ttu-id="c4139-118">Version</span><span class="sxs-lookup"><span data-stu-id="c4139-118">Version</span></span>](version.md)|<span data-ttu-id="c4139-119">x</span><span class="sxs-lookup"><span data-stu-id="c4139-119">x</span></span>|<span data-ttu-id="c4139-120">x</span><span class="sxs-lookup"><span data-stu-id="c4139-120">x</span></span>|<span data-ttu-id="c4139-121">x</span><span class="sxs-lookup"><span data-stu-id="c4139-121">x</span></span>|
|[<span data-ttu-id="c4139-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="c4139-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="c4139-123">x</span><span class="sxs-lookup"><span data-stu-id="c4139-123">x</span></span>|<span data-ttu-id="c4139-124">x</span><span class="sxs-lookup"><span data-stu-id="c4139-124">x</span></span>|<span data-ttu-id="c4139-125">x</span><span class="sxs-lookup"><span data-stu-id="c4139-125">x</span></span>|
|[<span data-ttu-id="c4139-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="c4139-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="c4139-127">x</span><span class="sxs-lookup"><span data-stu-id="c4139-127">x</span></span>|<span data-ttu-id="c4139-128">x</span><span class="sxs-lookup"><span data-stu-id="c4139-128">x</span></span>|<span data-ttu-id="c4139-129">x</span><span class="sxs-lookup"><span data-stu-id="c4139-129">x</span></span>|
|[<span data-ttu-id="c4139-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="c4139-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="c4139-131">x</span><span class="sxs-lookup"><span data-stu-id="c4139-131">x</span></span>||<span data-ttu-id="c4139-132">x</span><span class="sxs-lookup"><span data-stu-id="c4139-132">x</span></span>|
|[<span data-ttu-id="c4139-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="c4139-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="c4139-134">x</span><span class="sxs-lookup"><span data-stu-id="c4139-134">x</span></span>|<span data-ttu-id="c4139-135">x</span><span class="sxs-lookup"><span data-stu-id="c4139-135">x</span></span>|<span data-ttu-id="c4139-136">x</span><span class="sxs-lookup"><span data-stu-id="c4139-136">x</span></span>|
|[<span data-ttu-id="c4139-137">Description</span><span class="sxs-lookup"><span data-stu-id="c4139-137">Description</span></span>](description.md)|<span data-ttu-id="c4139-138">x</span><span class="sxs-lookup"><span data-stu-id="c4139-138">x</span></span>|<span data-ttu-id="c4139-139">x</span><span class="sxs-lookup"><span data-stu-id="c4139-139">x</span></span>|<span data-ttu-id="c4139-140">x</span><span class="sxs-lookup"><span data-stu-id="c4139-140">x</span></span>|
|[<span data-ttu-id="c4139-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="c4139-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="c4139-142">x</span><span class="sxs-lookup"><span data-stu-id="c4139-142">x</span></span>||
|[<span data-ttu-id="c4139-143">Permissions</span><span class="sxs-lookup"><span data-stu-id="c4139-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="c4139-144">x</span><span class="sxs-lookup"><span data-stu-id="c4139-144">x</span></span>||<span data-ttu-id="c4139-145">x</span><span class="sxs-lookup"><span data-stu-id="c4139-145">x</span></span>|
|[<span data-ttu-id="c4139-146">Règle</span><span class="sxs-lookup"><span data-stu-id="c4139-146">Rule</span></span>](rule.md)||<span data-ttu-id="c4139-147">x</span><span class="sxs-lookup"><span data-stu-id="c4139-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="c4139-148">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="c4139-148">Can contain</span></span>

|<span data-ttu-id="c4139-149">**Element**</span><span class="sxs-lookup"><span data-stu-id="c4139-149">**Element**</span></span>|<span data-ttu-id="c4139-150">**Content**</span><span class="sxs-lookup"><span data-stu-id="c4139-150">**Content**</span></span>|<span data-ttu-id="c4139-151">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="c4139-151">**Mail**</span></span>|<span data-ttu-id="c4139-152">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="c4139-152">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="c4139-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="c4139-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="c4139-154">x</span><span class="sxs-lookup"><span data-stu-id="c4139-154">x</span></span>|<span data-ttu-id="c4139-155">x</span><span class="sxs-lookup"><span data-stu-id="c4139-155">x</span></span>|<span data-ttu-id="c4139-156">x</span><span class="sxs-lookup"><span data-stu-id="c4139-156">x</span></span>|
|[<span data-ttu-id="c4139-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="c4139-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="c4139-158">x</span><span class="sxs-lookup"><span data-stu-id="c4139-158">x</span></span>|<span data-ttu-id="c4139-159">x</span><span class="sxs-lookup"><span data-stu-id="c4139-159">x</span></span>|<span data-ttu-id="c4139-160">x</span><span class="sxs-lookup"><span data-stu-id="c4139-160">x</span></span>|
|[<span data-ttu-id="c4139-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="c4139-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="c4139-162">x</span><span class="sxs-lookup"><span data-stu-id="c4139-162">x</span></span>|<span data-ttu-id="c4139-163">x</span><span class="sxs-lookup"><span data-stu-id="c4139-163">x</span></span>|<span data-ttu-id="c4139-164">x</span><span class="sxs-lookup"><span data-stu-id="c4139-164">x</span></span>|
|[<span data-ttu-id="c4139-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="c4139-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="c4139-166">x</span><span class="sxs-lookup"><span data-stu-id="c4139-166">x</span></span>|<span data-ttu-id="c4139-167">x</span><span class="sxs-lookup"><span data-stu-id="c4139-167">x</span></span>|<span data-ttu-id="c4139-168">x</span><span class="sxs-lookup"><span data-stu-id="c4139-168">x</span></span>|
|[<span data-ttu-id="c4139-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="c4139-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="c4139-170">x</span><span class="sxs-lookup"><span data-stu-id="c4139-170">x</span></span>|<span data-ttu-id="c4139-171">x</span><span class="sxs-lookup"><span data-stu-id="c4139-171">x</span></span>|<span data-ttu-id="c4139-172">x</span><span class="sxs-lookup"><span data-stu-id="c4139-172">x</span></span>|
|[<span data-ttu-id="c4139-173">Hôtes</span><span class="sxs-lookup"><span data-stu-id="c4139-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="c4139-174">x</span><span class="sxs-lookup"><span data-stu-id="c4139-174">x</span></span>|<span data-ttu-id="c4139-175">x</span><span class="sxs-lookup"><span data-stu-id="c4139-175">x</span></span>|<span data-ttu-id="c4139-176">x</span><span class="sxs-lookup"><span data-stu-id="c4139-176">x</span></span>|
|[<span data-ttu-id="c4139-177">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c4139-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="c4139-178">x</span><span class="sxs-lookup"><span data-stu-id="c4139-178">x</span></span>|<span data-ttu-id="c4139-179">x</span><span class="sxs-lookup"><span data-stu-id="c4139-179">x</span></span>|<span data-ttu-id="c4139-180">x</span><span class="sxs-lookup"><span data-stu-id="c4139-180">x</span></span>|
|[<span data-ttu-id="c4139-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="c4139-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="c4139-182">x</span><span class="sxs-lookup"><span data-stu-id="c4139-182">x</span></span>|||
|[<span data-ttu-id="c4139-183">Permissions</span><span class="sxs-lookup"><span data-stu-id="c4139-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="c4139-184">x</span><span class="sxs-lookup"><span data-stu-id="c4139-184">x</span></span>||
|[<span data-ttu-id="c4139-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="c4139-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="c4139-186">x</span><span class="sxs-lookup"><span data-stu-id="c4139-186">x</span></span>||
|[<span data-ttu-id="c4139-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="c4139-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="c4139-188">x</span><span class="sxs-lookup"><span data-stu-id="c4139-188">x</span></span>|
|[<span data-ttu-id="c4139-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="c4139-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="c4139-190">x</span><span class="sxs-lookup"><span data-stu-id="c4139-190">x</span></span>|<span data-ttu-id="c4139-191">x</span><span class="sxs-lookup"><span data-stu-id="c4139-191">x</span></span>|<span data-ttu-id="c4139-192">x</span><span class="sxs-lookup"><span data-stu-id="c4139-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="c4139-193">Attributs</span><span class="sxs-lookup"><span data-stu-id="c4139-193">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="c4139-194">xmlns</span><span class="sxs-lookup"><span data-stu-id="c4139-194">xmlns</span></span>|<span data-ttu-id="c4139-p101">Définit la version de schéma et l’espace de noms du manifeste de complément Office. Cet attribut doit toujours être défini sur `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="c4139-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="c4139-197">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="c4139-197">xmlns:xsi</span></span>|<span data-ttu-id="c4139-p102">Définit l’instance XMLSchema. Cet attribut doit toujours être défini sur `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="c4139-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="c4139-200">xsi:type</span><span class="sxs-lookup"><span data-stu-id="c4139-200">xsi:type</span></span>|<span data-ttu-id="c4139-p103">Définit le type de complément Office. Cet attribut doit être défini sur l’une des options suivantes : `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`.</span><span class="sxs-lookup"><span data-stu-id="c4139-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
