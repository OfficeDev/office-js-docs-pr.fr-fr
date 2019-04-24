---
title: Élément OfficeApp dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 86f38ab77e98bb01370e40c8ada38bae171e0c2d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450456"
---
# <a name="officeapp-element"></a><span data-ttu-id="a24a1-102">OfficeApp, élément</span><span class="sxs-lookup"><span data-stu-id="a24a1-102">OfficeApp element</span></span>

<span data-ttu-id="a24a1-103">Élément racine dans le manifeste d’un complément Office.</span><span class="sxs-lookup"><span data-stu-id="a24a1-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="a24a1-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="a24a1-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a24a1-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="a24a1-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="a24a1-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="a24a1-106">Contained in</span></span>

 <span data-ttu-id="a24a1-107">_none_</span><span class="sxs-lookup"><span data-stu-id="a24a1-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="a24a1-108">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="a24a1-108">Must contain</span></span>

|<span data-ttu-id="a24a1-109">**Élément**</span><span class="sxs-lookup"><span data-stu-id="a24a1-109">**Element**</span></span>|<span data-ttu-id="a24a1-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="a24a1-110">**Content**</span></span>|<span data-ttu-id="a24a1-111">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="a24a1-111">**Mail**</span></span>|<span data-ttu-id="a24a1-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="a24a1-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="a24a1-113">Id</span><span class="sxs-lookup"><span data-stu-id="a24a1-113">Id</span></span>](id.md)|<span data-ttu-id="a24a1-114">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-114">x</span></span>|<span data-ttu-id="a24a1-115">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-115">x</span></span>|<span data-ttu-id="a24a1-116">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-116">x</span></span>|
|[<span data-ttu-id="a24a1-117">Version</span><span class="sxs-lookup"><span data-stu-id="a24a1-117">Version</span></span>](version.md)|<span data-ttu-id="a24a1-118">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-118">x</span></span>|<span data-ttu-id="a24a1-119">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-119">x</span></span>|<span data-ttu-id="a24a1-120">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-120">x</span></span>|
|[<span data-ttu-id="a24a1-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="a24a1-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="a24a1-122">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-122">x</span></span>|<span data-ttu-id="a24a1-123">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-123">x</span></span>|<span data-ttu-id="a24a1-124">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-124">x</span></span>|
|[<span data-ttu-id="a24a1-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="a24a1-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="a24a1-126">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-126">x</span></span>|<span data-ttu-id="a24a1-127">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-127">x</span></span>|<span data-ttu-id="a24a1-128">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-128">x</span></span>|
|[<span data-ttu-id="a24a1-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="a24a1-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="a24a1-130">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-130">x</span></span>||<span data-ttu-id="a24a1-131">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-131">x</span></span>|
|[<span data-ttu-id="a24a1-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="a24a1-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="a24a1-133">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-133">x</span></span>|<span data-ttu-id="a24a1-134">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-134">x</span></span>|<span data-ttu-id="a24a1-135">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-135">x</span></span>|
|[<span data-ttu-id="a24a1-136">Description</span><span class="sxs-lookup"><span data-stu-id="a24a1-136">Description</span></span>](description.md)|<span data-ttu-id="a24a1-137">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-137">x</span></span>|<span data-ttu-id="a24a1-138">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-138">x</span></span>|<span data-ttu-id="a24a1-139">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-139">x</span></span>|
|[<span data-ttu-id="a24a1-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="a24a1-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="a24a1-141">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-141">x</span></span>||
|[<span data-ttu-id="a24a1-142">Permissions</span><span class="sxs-lookup"><span data-stu-id="a24a1-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="a24a1-143">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-143">x</span></span>||<span data-ttu-id="a24a1-144">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-144">x</span></span>|
|[<span data-ttu-id="a24a1-145">Règle</span><span class="sxs-lookup"><span data-stu-id="a24a1-145">Rule</span></span>](rule.md)||<span data-ttu-id="a24a1-146">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="a24a1-147">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="a24a1-147">Can contain</span></span>

|<span data-ttu-id="a24a1-148">**Element**</span><span class="sxs-lookup"><span data-stu-id="a24a1-148">**Element**</span></span>|<span data-ttu-id="a24a1-149">**Content**</span><span class="sxs-lookup"><span data-stu-id="a24a1-149">**Content**</span></span>|<span data-ttu-id="a24a1-150">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="a24a1-150">**Mail**</span></span>|<span data-ttu-id="a24a1-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="a24a1-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="a24a1-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="a24a1-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="a24a1-153">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-153">x</span></span>|<span data-ttu-id="a24a1-154">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-154">x</span></span>|<span data-ttu-id="a24a1-155">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-155">x</span></span>|
|[<span data-ttu-id="a24a1-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="a24a1-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="a24a1-157">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-157">x</span></span>|<span data-ttu-id="a24a1-158">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-158">x</span></span>|<span data-ttu-id="a24a1-159">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-159">x</span></span>|
|[<span data-ttu-id="a24a1-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="a24a1-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="a24a1-161">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-161">x</span></span>|<span data-ttu-id="a24a1-162">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-162">x</span></span>|<span data-ttu-id="a24a1-163">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-163">x</span></span>|
|[<span data-ttu-id="a24a1-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="a24a1-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="a24a1-165">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-165">x</span></span>|<span data-ttu-id="a24a1-166">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-166">x</span></span>|<span data-ttu-id="a24a1-167">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-167">x</span></span>|
|[<span data-ttu-id="a24a1-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="a24a1-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="a24a1-169">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-169">x</span></span>|<span data-ttu-id="a24a1-170">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-170">x</span></span>|<span data-ttu-id="a24a1-171">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-171">x</span></span>|
|[<span data-ttu-id="a24a1-172">Hôtes</span><span class="sxs-lookup"><span data-stu-id="a24a1-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="a24a1-173">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-173">x</span></span>|<span data-ttu-id="a24a1-174">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-174">x</span></span>|<span data-ttu-id="a24a1-175">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-175">x</span></span>|
|[<span data-ttu-id="a24a1-176">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a24a1-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="a24a1-177">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-177">x</span></span>|<span data-ttu-id="a24a1-178">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-178">x</span></span>|<span data-ttu-id="a24a1-179">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-179">x</span></span>|
|[<span data-ttu-id="a24a1-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="a24a1-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="a24a1-181">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-181">x</span></span>|||
|[<span data-ttu-id="a24a1-182">Permissions</span><span class="sxs-lookup"><span data-stu-id="a24a1-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="a24a1-183">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-183">x</span></span>||
|[<span data-ttu-id="a24a1-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="a24a1-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="a24a1-185">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-185">x</span></span>||
|[<span data-ttu-id="a24a1-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="a24a1-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="a24a1-187">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-187">x</span></span>|
|[<span data-ttu-id="a24a1-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="a24a1-188">VersionOverrides</span></span>](versionoverrides.md)||<span data-ttu-id="a24a1-189">x</span><span class="sxs-lookup"><span data-stu-id="a24a1-189">x</span></span>||

## <a name="attributes"></a><span data-ttu-id="a24a1-190">Attributs</span><span class="sxs-lookup"><span data-stu-id="a24a1-190">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="a24a1-191">xmlns</span><span class="sxs-lookup"><span data-stu-id="a24a1-191">xmlns</span></span>|<span data-ttu-id="a24a1-p101">Définit la version de schéma et l’espace de noms du manifeste de complément Office. Cet attribut doit toujours être défini sur `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="a24a1-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="a24a1-194">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="a24a1-194">xmlns:xsi</span></span>|<span data-ttu-id="a24a1-p102">Définit l’instance XMLSchema. Cet attribut doit toujours être défini sur `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="a24a1-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="a24a1-197">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a24a1-197">xsi:type</span></span>|<span data-ttu-id="a24a1-p103">Définit le type de complément Office. Cet attribut doit être défini sur l’une des options suivantes : `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`.</span><span class="sxs-lookup"><span data-stu-id="a24a1-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
