---
title: Manifeste XML des compléments Office
description: Obtenez une vue d’ensemble du manifeste de Complément Office et de ses applications.
ms.date: 03/18/2020
localization_priority: Priority
ms.openlocfilehash: 495638ee70630c5330e800419076463273bd2491
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293353"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="3c5ce-103">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="3c5ce-103">Office Add-ins XML manifest</span></span>

<span data-ttu-id="3c5ce-104">Le fichier manifeste XML d’un complément Office la manière dont votre complément doit être activé lorsqu’un utilisateur final l’installe et l’utilise avec des documents et des applications Office.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-104">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="3c5ce-105">Un fichier de manifeste XML basé sur ce schéma permet à un Complément Office d’effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="3c5ce-105">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="3c5ce-106">Se décrire en fournissant un ID, une version, une description, un nom d’affichage et un paramètre régional par défaut.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-106">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="3c5ce-107">Précisez les images utilisées pour l'image de marque du complément et l'iconographie utilisée pour [commandes complémentaires][] dans le ruban d'application de l'Office.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-107">Specify the images used for branding the add-in and iconography used for [add-in commands][] in the Office app ribbon.</span></span>

* <span data-ttu-id="3c5ce-108">Spécifier comment le complément s’intègre à Office, y compris les interfaces utilisateur personnalisées, telles que les boutons du ruban créés par le complément.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-108">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="3c5ce-109">Spécifier les dimensions par défaut demandées pour des compléments de contenu, et la hauteur demandée pour des compléments Outlook.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-109">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="3c5ce-110">Déclarer les autorisations que le Complément Office nécessite, par exemple la lecture du document ou l’écriture dans celui-ci.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-110">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="3c5ce-111">Pour des compléments Outlook, définir la ou les règles qui spécifient le contexte dans lequel ils seront activés et seront en interaction avec un message, un rendez-vous ou un élément de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-111">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="required-elements"></a><span data-ttu-id="3c5ce-112">Éléments requis</span><span class="sxs-lookup"><span data-stu-id="3c5ce-112">Required elements</span></span>

<span data-ttu-id="3c5ce-113">Le tableau suivant spécifie les éléments qui sont requis pour les trois types de compléments Office.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-113">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

> [!NOTE]
> <span data-ttu-id="3c5ce-114">Il existe également un ordre obligatoire d’apparition des éléments au sein de leur élément parent.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-114">There is also a mandatory order in which elements must appear within their parent element.</span></span> <span data-ttu-id="3c5ce-115">Pour plus d’informations, reportez-vous à la rubrique [Comment trouver l’ordre approprié d’éléments manifeste](manifest-element-ordering.md).</span><span class="sxs-lookup"><span data-stu-id="3c5ce-115">For more information see [How to find the proper order of manifest elements](manifest-element-ordering.md).</span></span>


### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="3c5ce-116">Éléments requis par type de complément Office</span><span class="sxs-lookup"><span data-stu-id="3c5ce-116">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="3c5ce-117">Élément</span><span class="sxs-lookup"><span data-stu-id="3c5ce-117">Element</span></span>                                                                                      | <span data-ttu-id="3c5ce-118">Contenu</span><span class="sxs-lookup"><span data-stu-id="3c5ce-118">Content</span></span> | <span data-ttu-id="3c5ce-119">Volet de tâches</span><span class="sxs-lookup"><span data-stu-id="3c5ce-119">Task pane</span></span> | <span data-ttu-id="3c5ce-120">Outlook</span><span class="sxs-lookup"><span data-stu-id="3c5ce-120">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="3c5ce-121">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-121">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="3c5ce-122">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-122">X</span></span>    |     <span data-ttu-id="3c5ce-123">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-123">X</span></span>     |    <span data-ttu-id="3c5ce-124">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-124">X</span></span>    |
| <span data-ttu-id="3c5ce-125">
  [Id][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-125">[Id][]</span></span>                                                                                       |    <span data-ttu-id="3c5ce-126">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-126">X</span></span>    |     <span data-ttu-id="3c5ce-127">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-127">X</span></span>     |    <span data-ttu-id="3c5ce-128">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-128">X</span></span>    |
| <span data-ttu-id="3c5ce-129">[Version][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-129">[Version][]</span></span>                                                                                  |    <span data-ttu-id="3c5ce-130">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-130">X</span></span>    |     <span data-ttu-id="3c5ce-131">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-131">X</span></span>     |    <span data-ttu-id="3c5ce-132">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-132">X</span></span>    |
| <span data-ttu-id="3c5ce-133">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-133">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="3c5ce-134">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-134">X</span></span>    |     <span data-ttu-id="3c5ce-135">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-135">X</span></span>     |    <span data-ttu-id="3c5ce-136">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-136">X</span></span>    |
| <span data-ttu-id="3c5ce-137">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-137">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="3c5ce-138">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-138">X</span></span>    |     <span data-ttu-id="3c5ce-139">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-139">X</span></span>     |    <span data-ttu-id="3c5ce-140">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-140">X</span></span>    |
| <span data-ttu-id="3c5ce-141">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-141">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="3c5ce-142">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-142">X</span></span>    |     <span data-ttu-id="3c5ce-143">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-143">X</span></span>     |    <span data-ttu-id="3c5ce-144">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-144">X</span></span>    |
| <span data-ttu-id="3c5ce-145">[Description][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-145">[Description][]</span></span>                                                                              |    <span data-ttu-id="3c5ce-146">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-146">X</span></span>    |     <span data-ttu-id="3c5ce-147">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-147">X</span></span>     |    <span data-ttu-id="3c5ce-148">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-148">X</span></span>    |
| <span data-ttu-id="3c5ce-149">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-149">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="3c5ce-150">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-150">X</span></span>    |     <span data-ttu-id="3c5ce-151">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-151">X</span></span>     |    <span data-ttu-id="3c5ce-152">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-152">X</span></span>    |
| <span data-ttu-id="3c5ce-153">[SupportUrl][]\*\*</span><span class="sxs-lookup"><span data-stu-id="3c5ce-153">[SupportUrl][]\*\*</span></span>                                                                           |    <span data-ttu-id="3c5ce-154">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-154">X</span></span>    |     <span data-ttu-id="3c5ce-155">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-155">X</span></span>     |    <span data-ttu-id="3c5ce-156">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-156">X</span></span>    |
| <span data-ttu-id="3c5ce-157">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-157">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="3c5ce-158">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-158">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="3c5ce-159">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-159">X</span></span>    |     <span data-ttu-id="3c5ce-160">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-160">X</span></span>     |         |
| <span data-ttu-id="3c5ce-161">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-161">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="3c5ce-162">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-162">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="3c5ce-163">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-163">X</span></span>    |     <span data-ttu-id="3c5ce-164">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-164">X</span></span>     |         |
| <span data-ttu-id="3c5ce-165">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-165">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="3c5ce-166">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-166">X</span></span>    |
| <span data-ttu-id="3c5ce-167">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-167">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="3c5ce-168">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-168">X</span></span>    |
| <span data-ttu-id="3c5ce-169">
  [Permissions (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-169">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="3c5ce-170">
  [Permissions (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-170">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="3c5ce-171">
  [Permissions (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-171">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="3c5ce-172">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-172">X</span></span>    |     <span data-ttu-id="3c5ce-173">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-173">X</span></span>     |    <span data-ttu-id="3c5ce-174">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-174">X</span></span>    |
| <span data-ttu-id="3c5ce-175">
  [Rule (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-175">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="3c5ce-176">
  [Rule (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-176">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="3c5ce-177">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-177">X</span></span>    |
| <span data-ttu-id="3c5ce-178">[Requirements (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-178">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="3c5ce-179">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-179">X</span></span>    |
| <span data-ttu-id="3c5ce-180">[Set\*][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-180">[Set\*][]</span></span><br/><span data-ttu-id="3c5ce-181">[Sets (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-181">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="3c5ce-182">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-182">X</span></span>    |
| <span data-ttu-id="3c5ce-183">[Form\*][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-183">[Form\*][]</span></span><br/><span data-ttu-id="3c5ce-184">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-184">[FormSettings\*][]</span></span>                                                              |         |           |    <span data-ttu-id="3c5ce-185">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-185">X</span></span>    |
| <span data-ttu-id="3c5ce-186">[Sets (Requirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-186">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="3c5ce-187">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-187">X</span></span>    |     <span data-ttu-id="3c5ce-188">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-188">X</span></span>     |         |
| <span data-ttu-id="3c5ce-189">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-189">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="3c5ce-190">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-190">X</span></span>    |     <span data-ttu-id="3c5ce-191">X</span><span class="sxs-lookup"><span data-stu-id="3c5ce-191">X</span></span>     |         |

<span data-ttu-id="3c5ce-192">_\*Ajouté dans le schéma de manifeste du complément Office version 1.1._</span><span class="sxs-lookup"><span data-stu-id="3c5ce-192">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<span data-ttu-id="3c5ce-193">_\*\* SupportUrl n’est obligatoire que pour les compléments distribués via AppSource._</span><span class="sxs-lookup"><span data-stu-id="3c5ce-193">_\*\* SupportUrl is only required for add-ins that are distributed through AppSource._</span></span>

<!-- Links for above table -->

[officeapp]: ../reference/manifest/officeapp.md
[id]: ../reference/manifest/id.md
[version]: ../reference/manifest/version.md
[providername]: ../reference/manifest/providername.md
[defaultlocale]: ../reference/manifest/defaultlocale.md
[displayname]: ../reference/manifest/displayname.md
[description]: ../reference/manifest/description.md
[iconurl]: ../reference/manifest/iconurl.md
[supporturl]: ../reference/manifest/supporturl.md
[defaultsettings (contentapp)]: ../reference/manifest/defaultsettings.md
[defaultsettings (taskpaneapp)]: ../reference/manifest/defaultsettings.md
[sourcelocation (contentapp)]: ../reference/manifest/sourcelocation.md
[sourcelocation (taskpaneapp)]: ../reference/manifest/sourcelocation.md
[desktopsettings]: /previous-versions/office/fp179684%28v=office.15%29
[sourcelocation (mailapp)]: /previous-versions/office/fp123668%28v=office.15%29
[permissions (contentapp)]: ../reference/manifest/permissions.md
[permissions (taskpaneapp)]: ../reference/manifest/permissions.md
[permissions (mailapp)]: ../reference/manifest/permissions.md
[rule (rulecollection)]: ../reference/manifest/rule.md
[rule (mailapp)]: ../reference/manifest/rule.md
[requirements (mailapp)*]: ../reference/manifest/requirements.md
[set*]: ../reference/manifest/set.md
[sets (mailapprequirements)*]: ../reference/manifest/sets.md
[form*]: ../reference/manifest/form.md
[formsettings*]: ../reference/manifest/formsettings.md
[sets (requirements)*]: ../reference/manifest/sets.md
[hosts*]: ../reference/manifest/hosts.md

## <a name="hosting-requirements"></a><span data-ttu-id="3c5ce-221">Configuration requise pour l’hébergement</span><span class="sxs-lookup"><span data-stu-id="3c5ce-221">Hosting requirements</span></span>

<span data-ttu-id="3c5ce-222">Tous les URI des images, tels que ceux utilisés pour les [commandes de complément][], doivent prendre en charge la mise en cache.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-222">All image URIs, such as those used for [add-in commands][], must support caching.</span></span> <span data-ttu-id="3c5ce-223">Le serveur qui héberge l’image ne doit pas renvoyer d’en-tête `Cache-Control` spécifiant `no-cache`, `no-store` ou des options similaires dans la réponse HTTP.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-223">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="3c5ce-224">Toutes les URL, telles que les emplacements des fichiers source spécifiés dans l’élément [SourceLocation](../reference/manifest/sourcelocation.md), doivent être **sécurisées par une protection SSL (HTTPS)**.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-224">All URLs, such as the source file locations specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="3c5ce-225">Bonnes pratiques pour l’envoi dans AppSource</span><span class="sxs-lookup"><span data-stu-id="3c5ce-225">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="3c5ce-p103">Vérifiez que l’ID du complément est un GUID valide et unique. Vous trouverez des outils de génération de GUID sur Internet pour vous aider à créer un GUID unique.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-p103">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="3c5ce-228">Les compléments envoyés à AppSource doivent également inclure l’élément [SupportUrl](../reference/manifest/supporturl.md).</span><span class="sxs-lookup"><span data-stu-id="3c5ce-228">Add-ins submitted to AppSource must also include the [SupportUrl](../reference/manifest/supporturl.md) element.</span></span> <span data-ttu-id="3c5ce-229">Pour plus d’informations, reportez-vous à [Stratégies de validation pour les applications et les compléments envoyés à AppSource](/legal/marketplace/certification-policies).</span><span class="sxs-lookup"><span data-stu-id="3c5ce-229">For more information, see [Validation policies for apps and add-ins submitted to AppSource](/legal/marketplace/certification-policies).</span></span>

<span data-ttu-id="3c5ce-230">Utilisez uniquement l’élément [AppDomains](../reference/manifest/appdomains.md) pour spécifier des domaines différents de celui spécifié dans l’élément [SourceLocation](../reference/manifest/sourcelocation.md) pour les scénarios d’authentification.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-230">Only use the [AppDomains](../reference/manifest/appdomains.md) element to specify domains other than the one specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="3c5ce-231">Spécifier les domaines que vous souhaitez ouvrir dans la fenêtre de complément</span><span class="sxs-lookup"><span data-stu-id="3c5ce-231">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="3c5ce-232">Quand vous exécutez Office sur le web, votre volet Office peut accéder à n’importe quelle URL.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-232">When running in Office on the web, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="3c5ce-233">Cependant, sur les plateformes de bureau, si votre complément tente d’accéder à une URL située dans un autre domaine que celui qui héberge la page de démarrage (comme indiqué dans l’élément [SourceLocation](../reference/manifest/sourcelocation.md) du fichier manifeste), cette URL s’ouvre dans une nouvelle fenêtre de navigateur en dehors du volet de complément de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-233">However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office application.</span></span>

<span data-ttu-id="3c5ce-234">Pour remplacer ce comportement (version de bureau d’Office), spécifiez chaque domaine à ouvrir dans la fenêtre de complément dans la liste des domaines spécifiés dans l’élément [AppDomains](../reference/manifest/appdomains.md) du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-234">To override this (desktop Office) behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](../reference/manifest/appdomains.md) element of the manifest file.</span></span> <span data-ttu-id="3c5ce-235">Si le complément tente d’accéder à une URL située dans un domaine figurant dans cette liste, il s’ouvre dans le volet Office d’Office sur le web et de la version de bureau d’Office.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-235">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both Office on the web and desktop.</span></span> <span data-ttu-id="3c5ce-236">S’il tente d’accéder à une URL qui ne figure pas dans la liste, dans la version de bureau d’Office, cette URL s’ouvre dans une nouvelle fenêtre de navigateur (en dehors du volet de complément).</span><span class="sxs-lookup"><span data-stu-id="3c5ce-236">If it tries to go to a URL that isn't in the list, then, in desktop Office, that URL opens in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="3c5ce-237">Il existe deux exceptions à ce comportement :</span><span class="sxs-lookup"><span data-stu-id="3c5ce-237">There are two exceptions to this behavior:</span></span>
>
> - <span data-ttu-id="3c5ce-238">Il s’applique uniquement au volet racine du complément.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-238">It applies only to the root pane of the add-in.</span></span> <span data-ttu-id="3c5ce-239">S’il existe un iframe incorporé dans la page de complément, l’iframe peut être dirigé vers n’importe quelle URL, qu’elle figure dans la liste des **AppDomains** ou non, y compris dans la version de bureau d’Office.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-239">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>
> - <span data-ttu-id="3c5ce-240">Lorsqu’une boîte de dialogue est ouverte avec l’API [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-), l’URL transmise à la méthode doit se trouver dans le même domaine que le complément, mais la boîte de dialogue peut ensuite être redirigée vers n’importe quelle URL, même si elle est répertoriée dans **AppDomains**, y compris dans la version de bureau d’Office.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-240">When a dialog is opened with the [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-) API, the URL that is passed to the method must be in the same domain as the add-in, but the dialog can then be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>

<span data-ttu-id="3c5ce-241">L’exemple de manifeste XML suivant héberge sa page de complément principale dans le domaine `https://www.contoso.com` comme indiqué dans l’élément **SourceLocation**.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-241">The following XML manifest example hosts its main add-in page in the `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="3c5ce-242">Il indique également le domaine `https://www.northwindtraders.com` dans un élément [AppDomain](../reference/manifest/appdomain.md) au sein de la liste d’éléments **AppDomains**.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-242">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](../reference/manifest/appdomain.md) element within the **AppDomains** element list.</span></span> <span data-ttu-id="3c5ce-243">Si le complément ouvre une page dans le domaine `www.northwindtraders.com`, cette page s’ouvre dans le volet de complément, y compris dans le bureau Office.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-243">If the add-in goes to a page in the `www.northwindtraders.com` domain, that page opens in the add-in pane, even in Office desktop.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a><span data-ttu-id="3c5ce-244">Spécifier les domaines à partir desquels les appels d’API Office.js sont effectués</span><span class="sxs-lookup"><span data-stu-id="3c5ce-244">Specify domains from which Office.js API calls are made</span></span>

<span data-ttu-id="3c5ce-245">Votre complément peut effectuer des appels d’API Office.js à partir du domaine référencé dans l’élément [SourceLocation](../reference/manifest/sourcelocation.md) du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-245">Your add-in can make Office.js API calls from the domain referenced in the [SourceLocation](../reference/manifest/sourcelocation.md) element of the manifest file.</span></span> <span data-ttu-id="3c5ce-246">Si votre complément comporte d’autres IFrames qui nécessitent un accès aux API Office.js, ajoutez le domaine de cette URL source à la liste spécifiée dans l’élément [AppDomains](../reference/manifest/appdomains.md) du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-246">If you have other IFrames within your add-in that need to access Office.js APIs, add the domain of that source URL to the list specified in the [AppDomains](../reference/manifest/appdomains.md) element of the manifest file.</span></span> <span data-ttu-id="3c5ce-247">Si un IFrame associé à une source qui ne figure pas dans la liste `AppDomains` tente d’effectuer un appel d’API Office.js, le complément reçoit une [erreur d’autorisation refusée](../reference/javascript-api-for-office-error-codes.md).</span><span class="sxs-lookup"><span data-stu-id="3c5ce-247">If an IFrame with a source not contained in the `AppDomains` list attempts to make an Office.js API call, then the add-in will receive a [permission denied error](../reference/javascript-api-for-office-error-codes.md).</span></span>

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="3c5ce-248">Exemples et schémas de fichier XML manifeste version 1.1</span><span class="sxs-lookup"><span data-stu-id="3c5ce-248">Manifest v1.1 XML file examples and schemas</span></span>

<span data-ttu-id="3c5ce-249">Les sections suivantes présentent des exemples de fichiers manifeste XML version 1.1 pour des compléments de contenu, de volet Office et Outlook.</span><span class="sxs-lookup"><span data-stu-id="3c5ce-249">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-pane"></a>[<span data-ttu-id="3c5ce-250">Volet Office</span><span class="sxs-lookup"><span data-stu-id="3c5ce-250">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="3c5ce-251">Schémas de manifeste de compléments</span><span class="sxs-lookup"><span data-stu-id="3c5ce-251">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
  <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office app ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
              <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                  <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                  <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                    <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                    <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon80" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
                    <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
                <!-- Menu example -->
                <Control xsi:type="Menu" id="Contoso.Menu">
                  <Label resid="Contoso.Dropdown.Label" />
                  <Supertip>
                    <Title resid="Contoso.Dropdown.Label" />
                    <Description resid="Contoso.Dropdown.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane1.Url" />
                      </Action>
                    </Item>

                    <Item id="Contoso.Menu.Item2">
                      <Label resid="Contoso.Item2.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item2.Label" />
                        <Description resid="Contoso.Item2.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane2.Url" />
                      </Action>
                    </Item>

                  </Items>
                </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon16" DefaultValue="https://myCDN/Images/Button16x16.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon32" DefaultValue="https://myCDN/Images/Button32x32.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon80" DefaultValue="https://myCDN/Images/Button80x80.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://i.imgur.com/qDujiX0.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# <a name="content"></a>[<span data-ttu-id="3c5ce-252">Content</span><span class="sxs-lookup"><span data-stu-id="3c5ce-252">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="3c5ce-253">Schémas de manifeste de compléments</span><span class="sxs-lookup"><span data-stu-id="3c5ce-253">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# <a name="mail"></a>[<span data-ttu-id="3c5ce-254">Application de messagerie</span><span class="sxs-lookup"><span data-stu-id="3c5ce-254">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="3c5ce-255">Schémas de manifeste de compléments</span><span class="sxs-lookup"><span data-stu-id="3c5ce-255">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="3c5ce-256">Valider un manifeste de complément Office</span><span class="sxs-lookup"><span data-stu-id="3c5ce-256">Validate an Office Add-in's manifest</span></span>

<span data-ttu-id="3c5ce-257">Pour plus d’informations sur la validation d’un manifeste par rapport à la [XSD (XML Schema Definition)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), voir [Valider le manifeste d’un complément Office](../testing/troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="3c5ce-257">For information about validating a manifest against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), see [Validate an Office Add-in's manifest](../testing/troubleshoot-manifest.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3c5ce-258">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3c5ce-258">See also</span></span>

* [<span data-ttu-id="3c5ce-259">Comment trouver l’ordre approprié d’éléments manifeste</span><span class="sxs-lookup"><span data-stu-id="3c5ce-259">How to find the proper order of manifest elements</span></span>](manifest-element-ordering.md)
* <span data-ttu-id="3c5ce-260">[Création de commandes de complément dans votre manifeste][commandes de complément]</span><span class="sxs-lookup"><span data-stu-id="3c5ce-260">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="3c5ce-261">Spécifier les exigences en matière d’applications Office et de l’API</span><span class="sxs-lookup"><span data-stu-id="3c5ce-261">Specify Office applications and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="3c5ce-262">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="3c5ce-262">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="3c5ce-263">Référence de schéma pour les manifestes des compléments Office</span><span class="sxs-lookup"><span data-stu-id="3c5ce-263">Schema reference for Office Add-ins manifests</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
* [<span data-ttu-id="3c5ce-264">Mettre à jour la version du manifeste et de l’API</span><span class="sxs-lookup"><span data-stu-id="3c5ce-264">Update API and manifest version</span></span>](update-your-javascript-api-for-office-and-manifest-schema-version.md)
* [<span data-ttu-id="3c5ce-265">Identifier un complément COM équivalent</span><span class="sxs-lookup"><span data-stu-id="3c5ce-265">Identify an equivalent COM add-in</span></span>](make-office-add-in-compatible-with-existing-com-add-in.md)
* [<span data-ttu-id="3c5ce-266">Demande d’autorisations d’utilisation de l’API dans des compléments</span><span class="sxs-lookup"><span data-stu-id="3c5ce-266">Requesting permissions for API use in add-ins</span></span>](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
* [<span data-ttu-id="3c5ce-267">Valider le manifeste d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="3c5ce-267">Validate an Office Add-in's manifest</span></span>](../testing/troubleshoot-manifest.md)

[commandes de complément]: create-addin-commands.md
[add-in commands]: create-addin-commands.md
