---
title: Élément OfficeApp dans le fichier manifeste
description: L’élément OfficeApp est l’élément racine d’un manifeste de complément Office.
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 038933f2d06ee5f485dbdb7dd7abdbd95fb97c7d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720595"
---
# <a name="officeapp-element"></a><span data-ttu-id="1b273-103">OfficeApp, élément</span><span class="sxs-lookup"><span data-stu-id="1b273-103">OfficeApp element</span></span>

<span data-ttu-id="1b273-104">Élément racine dans le manifeste d’un complément Office.</span><span class="sxs-lookup"><span data-stu-id="1b273-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="1b273-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="1b273-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1b273-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="1b273-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="1b273-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="1b273-107">Contained in</span></span>

 <span data-ttu-id="1b273-108">_none_</span><span class="sxs-lookup"><span data-stu-id="1b273-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="1b273-109">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="1b273-109">Must contain</span></span>

|<span data-ttu-id="1b273-110">**Élément**</span><span class="sxs-lookup"><span data-stu-id="1b273-110">**Element**</span></span>|<span data-ttu-id="1b273-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="1b273-111">**Content**</span></span>|<span data-ttu-id="1b273-112">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="1b273-112">**Mail**</span></span>|<span data-ttu-id="1b273-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="1b273-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="1b273-114">Id</span><span class="sxs-lookup"><span data-stu-id="1b273-114">Id</span></span>](id.md)|<span data-ttu-id="1b273-115">x</span><span class="sxs-lookup"><span data-stu-id="1b273-115">x</span></span>|<span data-ttu-id="1b273-116">x</span><span class="sxs-lookup"><span data-stu-id="1b273-116">x</span></span>|<span data-ttu-id="1b273-117">x</span><span class="sxs-lookup"><span data-stu-id="1b273-117">x</span></span>|
|[<span data-ttu-id="1b273-118">Version</span><span class="sxs-lookup"><span data-stu-id="1b273-118">Version</span></span>](version.md)|<span data-ttu-id="1b273-119">x</span><span class="sxs-lookup"><span data-stu-id="1b273-119">x</span></span>|<span data-ttu-id="1b273-120">x</span><span class="sxs-lookup"><span data-stu-id="1b273-120">x</span></span>|<span data-ttu-id="1b273-121">x</span><span class="sxs-lookup"><span data-stu-id="1b273-121">x</span></span>|
|[<span data-ttu-id="1b273-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="1b273-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="1b273-123">x</span><span class="sxs-lookup"><span data-stu-id="1b273-123">x</span></span>|<span data-ttu-id="1b273-124">x</span><span class="sxs-lookup"><span data-stu-id="1b273-124">x</span></span>|<span data-ttu-id="1b273-125">x</span><span class="sxs-lookup"><span data-stu-id="1b273-125">x</span></span>|
|[<span data-ttu-id="1b273-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="1b273-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="1b273-127">x</span><span class="sxs-lookup"><span data-stu-id="1b273-127">x</span></span>|<span data-ttu-id="1b273-128">x</span><span class="sxs-lookup"><span data-stu-id="1b273-128">x</span></span>|<span data-ttu-id="1b273-129">x</span><span class="sxs-lookup"><span data-stu-id="1b273-129">x</span></span>|
|[<span data-ttu-id="1b273-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="1b273-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="1b273-131">x</span><span class="sxs-lookup"><span data-stu-id="1b273-131">x</span></span>||<span data-ttu-id="1b273-132">x</span><span class="sxs-lookup"><span data-stu-id="1b273-132">x</span></span>|
|[<span data-ttu-id="1b273-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="1b273-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="1b273-134">x</span><span class="sxs-lookup"><span data-stu-id="1b273-134">x</span></span>|<span data-ttu-id="1b273-135">x</span><span class="sxs-lookup"><span data-stu-id="1b273-135">x</span></span>|<span data-ttu-id="1b273-136">x</span><span class="sxs-lookup"><span data-stu-id="1b273-136">x</span></span>|
|[<span data-ttu-id="1b273-137">Description</span><span class="sxs-lookup"><span data-stu-id="1b273-137">Description</span></span>](description.md)|<span data-ttu-id="1b273-138">x</span><span class="sxs-lookup"><span data-stu-id="1b273-138">x</span></span>|<span data-ttu-id="1b273-139">x</span><span class="sxs-lookup"><span data-stu-id="1b273-139">x</span></span>|<span data-ttu-id="1b273-140">x</span><span class="sxs-lookup"><span data-stu-id="1b273-140">x</span></span>|
|[<span data-ttu-id="1b273-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="1b273-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="1b273-142">x</span><span class="sxs-lookup"><span data-stu-id="1b273-142">x</span></span>||
|[<span data-ttu-id="1b273-143">Permissions</span><span class="sxs-lookup"><span data-stu-id="1b273-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="1b273-144">x</span><span class="sxs-lookup"><span data-stu-id="1b273-144">x</span></span>||<span data-ttu-id="1b273-145">x</span><span class="sxs-lookup"><span data-stu-id="1b273-145">x</span></span>|
|[<span data-ttu-id="1b273-146">Règle</span><span class="sxs-lookup"><span data-stu-id="1b273-146">Rule</span></span>](rule.md)||<span data-ttu-id="1b273-147">x</span><span class="sxs-lookup"><span data-stu-id="1b273-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="1b273-148">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="1b273-148">Can contain</span></span>

|<span data-ttu-id="1b273-149">**Element**</span><span class="sxs-lookup"><span data-stu-id="1b273-149">**Element**</span></span>|<span data-ttu-id="1b273-150">**Content**</span><span class="sxs-lookup"><span data-stu-id="1b273-150">**Content**</span></span>|<span data-ttu-id="1b273-151">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="1b273-151">**Mail**</span></span>|<span data-ttu-id="1b273-152">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="1b273-152">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="1b273-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="1b273-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="1b273-154">x</span><span class="sxs-lookup"><span data-stu-id="1b273-154">x</span></span>|<span data-ttu-id="1b273-155">x</span><span class="sxs-lookup"><span data-stu-id="1b273-155">x</span></span>|<span data-ttu-id="1b273-156">x</span><span class="sxs-lookup"><span data-stu-id="1b273-156">x</span></span>|
|[<span data-ttu-id="1b273-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="1b273-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="1b273-158">x</span><span class="sxs-lookup"><span data-stu-id="1b273-158">x</span></span>|<span data-ttu-id="1b273-159">x</span><span class="sxs-lookup"><span data-stu-id="1b273-159">x</span></span>|<span data-ttu-id="1b273-160">x</span><span class="sxs-lookup"><span data-stu-id="1b273-160">x</span></span>|
|[<span data-ttu-id="1b273-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="1b273-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="1b273-162">x</span><span class="sxs-lookup"><span data-stu-id="1b273-162">x</span></span>|<span data-ttu-id="1b273-163">x</span><span class="sxs-lookup"><span data-stu-id="1b273-163">x</span></span>|<span data-ttu-id="1b273-164">x</span><span class="sxs-lookup"><span data-stu-id="1b273-164">x</span></span>|
|[<span data-ttu-id="1b273-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="1b273-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="1b273-166">x</span><span class="sxs-lookup"><span data-stu-id="1b273-166">x</span></span>|<span data-ttu-id="1b273-167">x</span><span class="sxs-lookup"><span data-stu-id="1b273-167">x</span></span>|<span data-ttu-id="1b273-168">x</span><span class="sxs-lookup"><span data-stu-id="1b273-168">x</span></span>|
|[<span data-ttu-id="1b273-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="1b273-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="1b273-170">x</span><span class="sxs-lookup"><span data-stu-id="1b273-170">x</span></span>|<span data-ttu-id="1b273-171">x</span><span class="sxs-lookup"><span data-stu-id="1b273-171">x</span></span>|<span data-ttu-id="1b273-172">x</span><span class="sxs-lookup"><span data-stu-id="1b273-172">x</span></span>|
|[<span data-ttu-id="1b273-173">Hôtes</span><span class="sxs-lookup"><span data-stu-id="1b273-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="1b273-174">x</span><span class="sxs-lookup"><span data-stu-id="1b273-174">x</span></span>|<span data-ttu-id="1b273-175">x</span><span class="sxs-lookup"><span data-stu-id="1b273-175">x</span></span>|<span data-ttu-id="1b273-176">x</span><span class="sxs-lookup"><span data-stu-id="1b273-176">x</span></span>|
|[<span data-ttu-id="1b273-177">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1b273-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="1b273-178">x</span><span class="sxs-lookup"><span data-stu-id="1b273-178">x</span></span>|<span data-ttu-id="1b273-179">x</span><span class="sxs-lookup"><span data-stu-id="1b273-179">x</span></span>|<span data-ttu-id="1b273-180">x</span><span class="sxs-lookup"><span data-stu-id="1b273-180">x</span></span>|
|[<span data-ttu-id="1b273-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="1b273-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="1b273-182">x</span><span class="sxs-lookup"><span data-stu-id="1b273-182">x</span></span>|||
|[<span data-ttu-id="1b273-183">Permissions</span><span class="sxs-lookup"><span data-stu-id="1b273-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="1b273-184">x</span><span class="sxs-lookup"><span data-stu-id="1b273-184">x</span></span>||
|[<span data-ttu-id="1b273-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="1b273-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="1b273-186">x</span><span class="sxs-lookup"><span data-stu-id="1b273-186">x</span></span>||
|[<span data-ttu-id="1b273-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="1b273-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="1b273-188">x</span><span class="sxs-lookup"><span data-stu-id="1b273-188">x</span></span>|
|[<span data-ttu-id="1b273-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="1b273-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="1b273-190">x</span><span class="sxs-lookup"><span data-stu-id="1b273-190">x</span></span>|<span data-ttu-id="1b273-191">x</span><span class="sxs-lookup"><span data-stu-id="1b273-191">x</span></span>|<span data-ttu-id="1b273-192">x</span><span class="sxs-lookup"><span data-stu-id="1b273-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="1b273-193">Attributs</span><span class="sxs-lookup"><span data-stu-id="1b273-193">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="1b273-194">xmlns</span><span class="sxs-lookup"><span data-stu-id="1b273-194">xmlns</span></span>|<span data-ttu-id="1b273-p101">Définit la version de schéma et l’espace de noms du manifeste de complément Office. Cet attribut doit toujours être défini sur `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="1b273-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="1b273-197">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="1b273-197">xmlns:xsi</span></span>|<span data-ttu-id="1b273-p102">Définit l’instance XMLSchema. Cet attribut doit toujours être défini sur `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="1b273-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="1b273-200">xsi:type</span><span class="sxs-lookup"><span data-stu-id="1b273-200">xsi:type</span></span>|<span data-ttu-id="1b273-p103">Définit le type de complément Office. Cet attribut doit être défini sur l’une des options suivantes : `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`.</span><span class="sxs-lookup"><span data-stu-id="1b273-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
