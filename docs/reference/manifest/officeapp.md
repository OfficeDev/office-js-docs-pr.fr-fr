---
title: Élément OfficeApp dans le fichier manifeste
description: ''
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 080025e62a56421dff942792f99ee672ce1db69a
ms.sourcegitcommit: c1dbea577ae6183523fb663d364422d2adbc8bcf
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/05/2020
ms.locfileid: "41773578"
---
# <a name="officeapp-element"></a><span data-ttu-id="472f7-102">OfficeApp, élément</span><span class="sxs-lookup"><span data-stu-id="472f7-102">OfficeApp element</span></span>

<span data-ttu-id="472f7-103">Élément racine dans le manifeste d’un complément Office.</span><span class="sxs-lookup"><span data-stu-id="472f7-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="472f7-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="472f7-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="472f7-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="472f7-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="472f7-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="472f7-106">Contained in</span></span>

 <span data-ttu-id="472f7-107">_none_</span><span class="sxs-lookup"><span data-stu-id="472f7-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="472f7-108">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="472f7-108">Must contain</span></span>

|<span data-ttu-id="472f7-109">**Élément**</span><span class="sxs-lookup"><span data-stu-id="472f7-109">**Element**</span></span>|<span data-ttu-id="472f7-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="472f7-110">**Content**</span></span>|<span data-ttu-id="472f7-111">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="472f7-111">**Mail**</span></span>|<span data-ttu-id="472f7-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="472f7-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="472f7-113">Id</span><span class="sxs-lookup"><span data-stu-id="472f7-113">Id</span></span>](id.md)|<span data-ttu-id="472f7-114">x</span><span class="sxs-lookup"><span data-stu-id="472f7-114">x</span></span>|<span data-ttu-id="472f7-115">x</span><span class="sxs-lookup"><span data-stu-id="472f7-115">x</span></span>|<span data-ttu-id="472f7-116">x</span><span class="sxs-lookup"><span data-stu-id="472f7-116">x</span></span>|
|[<span data-ttu-id="472f7-117">Version</span><span class="sxs-lookup"><span data-stu-id="472f7-117">Version</span></span>](version.md)|<span data-ttu-id="472f7-118">x</span><span class="sxs-lookup"><span data-stu-id="472f7-118">x</span></span>|<span data-ttu-id="472f7-119">x</span><span class="sxs-lookup"><span data-stu-id="472f7-119">x</span></span>|<span data-ttu-id="472f7-120">x</span><span class="sxs-lookup"><span data-stu-id="472f7-120">x</span></span>|
|[<span data-ttu-id="472f7-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="472f7-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="472f7-122">x</span><span class="sxs-lookup"><span data-stu-id="472f7-122">x</span></span>|<span data-ttu-id="472f7-123">x</span><span class="sxs-lookup"><span data-stu-id="472f7-123">x</span></span>|<span data-ttu-id="472f7-124">x</span><span class="sxs-lookup"><span data-stu-id="472f7-124">x</span></span>|
|[<span data-ttu-id="472f7-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="472f7-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="472f7-126">x</span><span class="sxs-lookup"><span data-stu-id="472f7-126">x</span></span>|<span data-ttu-id="472f7-127">x</span><span class="sxs-lookup"><span data-stu-id="472f7-127">x</span></span>|<span data-ttu-id="472f7-128">x</span><span class="sxs-lookup"><span data-stu-id="472f7-128">x</span></span>|
|[<span data-ttu-id="472f7-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="472f7-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="472f7-130">x</span><span class="sxs-lookup"><span data-stu-id="472f7-130">x</span></span>||<span data-ttu-id="472f7-131">x</span><span class="sxs-lookup"><span data-stu-id="472f7-131">x</span></span>|
|[<span data-ttu-id="472f7-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="472f7-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="472f7-133">x</span><span class="sxs-lookup"><span data-stu-id="472f7-133">x</span></span>|<span data-ttu-id="472f7-134">x</span><span class="sxs-lookup"><span data-stu-id="472f7-134">x</span></span>|<span data-ttu-id="472f7-135">x</span><span class="sxs-lookup"><span data-stu-id="472f7-135">x</span></span>|
|[<span data-ttu-id="472f7-136">Description</span><span class="sxs-lookup"><span data-stu-id="472f7-136">Description</span></span>](description.md)|<span data-ttu-id="472f7-137">x</span><span class="sxs-lookup"><span data-stu-id="472f7-137">x</span></span>|<span data-ttu-id="472f7-138">x</span><span class="sxs-lookup"><span data-stu-id="472f7-138">x</span></span>|<span data-ttu-id="472f7-139">x</span><span class="sxs-lookup"><span data-stu-id="472f7-139">x</span></span>|
|[<span data-ttu-id="472f7-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="472f7-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="472f7-141">x</span><span class="sxs-lookup"><span data-stu-id="472f7-141">x</span></span>||
|[<span data-ttu-id="472f7-142">Permissions</span><span class="sxs-lookup"><span data-stu-id="472f7-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="472f7-143">x</span><span class="sxs-lookup"><span data-stu-id="472f7-143">x</span></span>||<span data-ttu-id="472f7-144">x</span><span class="sxs-lookup"><span data-stu-id="472f7-144">x</span></span>|
|[<span data-ttu-id="472f7-145">Règle</span><span class="sxs-lookup"><span data-stu-id="472f7-145">Rule</span></span>](rule.md)||<span data-ttu-id="472f7-146">x</span><span class="sxs-lookup"><span data-stu-id="472f7-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="472f7-147">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="472f7-147">Can contain</span></span>

|<span data-ttu-id="472f7-148">**Element**</span><span class="sxs-lookup"><span data-stu-id="472f7-148">**Element**</span></span>|<span data-ttu-id="472f7-149">**Content**</span><span class="sxs-lookup"><span data-stu-id="472f7-149">**Content**</span></span>|<span data-ttu-id="472f7-150">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="472f7-150">**Mail**</span></span>|<span data-ttu-id="472f7-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="472f7-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="472f7-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="472f7-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="472f7-153">x</span><span class="sxs-lookup"><span data-stu-id="472f7-153">x</span></span>|<span data-ttu-id="472f7-154">x</span><span class="sxs-lookup"><span data-stu-id="472f7-154">x</span></span>|<span data-ttu-id="472f7-155">x</span><span class="sxs-lookup"><span data-stu-id="472f7-155">x</span></span>|
|[<span data-ttu-id="472f7-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="472f7-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="472f7-157">x</span><span class="sxs-lookup"><span data-stu-id="472f7-157">x</span></span>|<span data-ttu-id="472f7-158">x</span><span class="sxs-lookup"><span data-stu-id="472f7-158">x</span></span>|<span data-ttu-id="472f7-159">x</span><span class="sxs-lookup"><span data-stu-id="472f7-159">x</span></span>|
|[<span data-ttu-id="472f7-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="472f7-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="472f7-161">x</span><span class="sxs-lookup"><span data-stu-id="472f7-161">x</span></span>|<span data-ttu-id="472f7-162">x</span><span class="sxs-lookup"><span data-stu-id="472f7-162">x</span></span>|<span data-ttu-id="472f7-163">x</span><span class="sxs-lookup"><span data-stu-id="472f7-163">x</span></span>|
|[<span data-ttu-id="472f7-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="472f7-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="472f7-165">x</span><span class="sxs-lookup"><span data-stu-id="472f7-165">x</span></span>|<span data-ttu-id="472f7-166">x</span><span class="sxs-lookup"><span data-stu-id="472f7-166">x</span></span>|<span data-ttu-id="472f7-167">x</span><span class="sxs-lookup"><span data-stu-id="472f7-167">x</span></span>|
|[<span data-ttu-id="472f7-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="472f7-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="472f7-169">x</span><span class="sxs-lookup"><span data-stu-id="472f7-169">x</span></span>|<span data-ttu-id="472f7-170">x</span><span class="sxs-lookup"><span data-stu-id="472f7-170">x</span></span>|<span data-ttu-id="472f7-171">x</span><span class="sxs-lookup"><span data-stu-id="472f7-171">x</span></span>|
|[<span data-ttu-id="472f7-172">Hôtes</span><span class="sxs-lookup"><span data-stu-id="472f7-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="472f7-173">x</span><span class="sxs-lookup"><span data-stu-id="472f7-173">x</span></span>|<span data-ttu-id="472f7-174">x</span><span class="sxs-lookup"><span data-stu-id="472f7-174">x</span></span>|<span data-ttu-id="472f7-175">x</span><span class="sxs-lookup"><span data-stu-id="472f7-175">x</span></span>|
|[<span data-ttu-id="472f7-176">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="472f7-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="472f7-177">x</span><span class="sxs-lookup"><span data-stu-id="472f7-177">x</span></span>|<span data-ttu-id="472f7-178">x</span><span class="sxs-lookup"><span data-stu-id="472f7-178">x</span></span>|<span data-ttu-id="472f7-179">x</span><span class="sxs-lookup"><span data-stu-id="472f7-179">x</span></span>|
|[<span data-ttu-id="472f7-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="472f7-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="472f7-181">x</span><span class="sxs-lookup"><span data-stu-id="472f7-181">x</span></span>|||
|[<span data-ttu-id="472f7-182">Permissions</span><span class="sxs-lookup"><span data-stu-id="472f7-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="472f7-183">x</span><span class="sxs-lookup"><span data-stu-id="472f7-183">x</span></span>||
|[<span data-ttu-id="472f7-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="472f7-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="472f7-185">x</span><span class="sxs-lookup"><span data-stu-id="472f7-185">x</span></span>||
|[<span data-ttu-id="472f7-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="472f7-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="472f7-187">x</span><span class="sxs-lookup"><span data-stu-id="472f7-187">x</span></span>|
|[<span data-ttu-id="472f7-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="472f7-188">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="472f7-189">x</span><span class="sxs-lookup"><span data-stu-id="472f7-189">x</span></span>|<span data-ttu-id="472f7-190">x</span><span class="sxs-lookup"><span data-stu-id="472f7-190">x</span></span>|<span data-ttu-id="472f7-191">x</span><span class="sxs-lookup"><span data-stu-id="472f7-191">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="472f7-192">Attributs</span><span class="sxs-lookup"><span data-stu-id="472f7-192">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="472f7-193">xmlns</span><span class="sxs-lookup"><span data-stu-id="472f7-193">xmlns</span></span>|<span data-ttu-id="472f7-p101">Définit la version de schéma et l’espace de noms du manifeste de complément Office. Cet attribut doit toujours être défini sur `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="472f7-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="472f7-196">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="472f7-196">xmlns:xsi</span></span>|<span data-ttu-id="472f7-p102">Définit l’instance XMLSchema. Cet attribut doit toujours être défini sur `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="472f7-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="472f7-199">xsi:type</span><span class="sxs-lookup"><span data-stu-id="472f7-199">xsi:type</span></span>|<span data-ttu-id="472f7-p103">Définit le type de complément Office. Cet attribut doit être défini sur l’une des options suivantes : `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`.</span><span class="sxs-lookup"><span data-stu-id="472f7-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
