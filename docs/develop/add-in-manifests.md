---
title: Manifeste XML des compléments Office
description: ''
ms.date: 02/09/2018
ms.openlocfilehash: 8d8363b80b948f30e13ccd8620178268e03f1d57
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505866"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="806c8-102">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="806c8-102">Office Add-ins XML manifest</span></span>

<span data-ttu-id="806c8-103">Le fichier manifeste XML d’un complément Office la manière dont votre complément doit être activé lorsqu’un utilisateur final l’installe et l’utilise avec des documents et des applications Office.</span><span class="sxs-lookup"><span data-stu-id="806c8-103">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="806c8-104">Un fichier de manifeste XML basé sur ce schéma permet à un Complément Office d’effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="806c8-104">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="806c8-105">Se décrire en fournissant un ID, une version, une description, un nom d’affichage et un paramètre régional par défaut.</span><span class="sxs-lookup"><span data-stu-id="806c8-105">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="806c8-106">Spécifier les images utilisées pour personnaliser le complément et l’iconographie utilisés pour les [commandes de complément][] dans le ruban Office.</span><span class="sxs-lookup"><span data-stu-id="806c8-106">Specify the images used for branding the Add-in and iconography used for [Add-in Commands][] in the Office Ribbon.</span></span>

* <span data-ttu-id="806c8-107">Spécifier comment le complément s’intègre à Office, y compris les interfaces utilisateur personnalisées, telles que les boutons du ruban créés par le complément.</span><span class="sxs-lookup"><span data-stu-id="806c8-107">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="806c8-108">Spécifier les dimensions par défaut demandées pour des compléments de contenu, et la hauteur demandée pour des compléments Outlook.</span><span class="sxs-lookup"><span data-stu-id="806c8-108">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="806c8-109">Déclarer les autorisations que le Complément Office nécessite, par exemple la lecture du document ou l’écriture dans celui-ci.</span><span class="sxs-lookup"><span data-stu-id="806c8-109">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="806c8-110">Pour des compléments Outlook, définir la ou les règles qui spécifient le contexte dans lequel ils seront activés et seront en interaction avec un message, un rendez-vous ou un élément de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="806c8-110">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

> [!NOTE]
> <span data-ttu-id="806c8-p101">Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de validation AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="806c8-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="required-elements"></a><span data-ttu-id="806c8-113">Éléments requis</span><span class="sxs-lookup"><span data-stu-id="806c8-113">Required elements</span></span>

<span data-ttu-id="806c8-114">Le tableau suivant spécifie les éléments qui sont requis pour les trois types de compléments Office.</span><span class="sxs-lookup"><span data-stu-id="806c8-114">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="806c8-115">Éléments requis par type de complément Office</span><span class="sxs-lookup"><span data-stu-id="806c8-115">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="806c8-116">Élément</span><span class="sxs-lookup"><span data-stu-id="806c8-116">Element</span></span>                                                                                      | <span data-ttu-id="806c8-117">Contenu</span><span class="sxs-lookup"><span data-stu-id="806c8-117">Content</span></span> | <span data-ttu-id="806c8-118">Volet Office</span><span class="sxs-lookup"><span data-stu-id="806c8-118">Task pane</span></span> | <span data-ttu-id="806c8-119">Outlook</span><span class="sxs-lookup"><span data-stu-id="806c8-119">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="806c8-120">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="806c8-120">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="806c8-121">X</span><span class="sxs-lookup"><span data-stu-id="806c8-121">X</span></span>    |     <span data-ttu-id="806c8-122">X</span><span class="sxs-lookup"><span data-stu-id="806c8-122">X</span></span>     |    <span data-ttu-id="806c8-123">X</span><span class="sxs-lookup"><span data-stu-id="806c8-123">X</span></span>    |
| <span data-ttu-id="806c8-124">[Id][]</span><span class="sxs-lookup"><span data-stu-id="806c8-124">[Id][]</span></span>                                                                                       |    <span data-ttu-id="806c8-125">X</span><span class="sxs-lookup"><span data-stu-id="806c8-125">X</span></span>    |     <span data-ttu-id="806c8-126">X</span><span class="sxs-lookup"><span data-stu-id="806c8-126">X</span></span>     |    <span data-ttu-id="806c8-127">X</span><span class="sxs-lookup"><span data-stu-id="806c8-127">X</span></span>    |
| <span data-ttu-id="806c8-128">[Version][]</span><span class="sxs-lookup"><span data-stu-id="806c8-128">[Version][]</span></span>                                                                                  |    <span data-ttu-id="806c8-129">X</span><span class="sxs-lookup"><span data-stu-id="806c8-129">X</span></span>    |     <span data-ttu-id="806c8-130">X</span><span class="sxs-lookup"><span data-stu-id="806c8-130">X</span></span>     |    <span data-ttu-id="806c8-131">X</span><span class="sxs-lookup"><span data-stu-id="806c8-131">X</span></span>    |
| <span data-ttu-id="806c8-132">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="806c8-132">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="806c8-133">X</span><span class="sxs-lookup"><span data-stu-id="806c8-133">X</span></span>    |     <span data-ttu-id="806c8-134">X</span><span class="sxs-lookup"><span data-stu-id="806c8-134">X</span></span>     |    <span data-ttu-id="806c8-135">X</span><span class="sxs-lookup"><span data-stu-id="806c8-135">X</span></span>    |
| <span data-ttu-id="806c8-136">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="806c8-136">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="806c8-137">X</span><span class="sxs-lookup"><span data-stu-id="806c8-137">X</span></span>    |     <span data-ttu-id="806c8-138">X</span><span class="sxs-lookup"><span data-stu-id="806c8-138">X</span></span>     |    <span data-ttu-id="806c8-139">X</span><span class="sxs-lookup"><span data-stu-id="806c8-139">X</span></span>    |
| <span data-ttu-id="806c8-140">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="806c8-140">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="806c8-141">X</span><span class="sxs-lookup"><span data-stu-id="806c8-141">X</span></span>    |     <span data-ttu-id="806c8-142">X</span><span class="sxs-lookup"><span data-stu-id="806c8-142">X</span></span>     |    <span data-ttu-id="806c8-143">X</span><span class="sxs-lookup"><span data-stu-id="806c8-143">X</span></span>    |
| <span data-ttu-id="806c8-144">[Description][]</span><span class="sxs-lookup"><span data-stu-id="806c8-144">[Description][]</span></span>                                                                              |    <span data-ttu-id="806c8-145">X</span><span class="sxs-lookup"><span data-stu-id="806c8-145">X</span></span>    |     <span data-ttu-id="806c8-146">X</span><span class="sxs-lookup"><span data-stu-id="806c8-146">X</span></span>     |    <span data-ttu-id="806c8-147">X</span><span class="sxs-lookup"><span data-stu-id="806c8-147">X</span></span>    |
| <span data-ttu-id="806c8-148">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="806c8-148">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="806c8-149">X</span><span class="sxs-lookup"><span data-stu-id="806c8-149">X</span></span>    |     <span data-ttu-id="806c8-150">X</span><span class="sxs-lookup"><span data-stu-id="806c8-150">X</span></span>     |    <span data-ttu-id="806c8-151">X</span><span class="sxs-lookup"><span data-stu-id="806c8-151">X</span></span>    |
| <span data-ttu-id="806c8-152">[HighResolutionIconUrl][]</span><span class="sxs-lookup"><span data-stu-id="806c8-152">[HighResolutionIconUrl][]</span></span>                                                                    |    <span data-ttu-id="806c8-153">X</span><span class="sxs-lookup"><span data-stu-id="806c8-153">X</span></span>    |     <span data-ttu-id="806c8-154">X</span><span class="sxs-lookup"><span data-stu-id="806c8-154">X</span></span>     |    <span data-ttu-id="806c8-155">X</span><span class="sxs-lookup"><span data-stu-id="806c8-155">X</span></span>    |
| <span data-ttu-id="806c8-156">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-156">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="806c8-157">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-157">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="806c8-158">X</span><span class="sxs-lookup"><span data-stu-id="806c8-158">X</span></span>    |     <span data-ttu-id="806c8-159">X</span><span class="sxs-lookup"><span data-stu-id="806c8-159">X</span></span>     |         |
| <span data-ttu-id="806c8-160">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-160">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="806c8-161">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-161">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="806c8-162">X</span><span class="sxs-lookup"><span data-stu-id="806c8-162">X</span></span>    |     <span data-ttu-id="806c8-163">X</span><span class="sxs-lookup"><span data-stu-id="806c8-163">X</span></span>     |         |
| <span data-ttu-id="806c8-164">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="806c8-164">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="806c8-165">X</span><span class="sxs-lookup"><span data-stu-id="806c8-165">X</span></span>    |
| <span data-ttu-id="806c8-166">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-166">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="806c8-167">X</span><span class="sxs-lookup"><span data-stu-id="806c8-167">X</span></span>    |
| <span data-ttu-id="806c8-168">[Permissions (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-168">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="806c8-169">[Permissions (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-169">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="806c8-170">[Permissions (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-170">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="806c8-171">X</span><span class="sxs-lookup"><span data-stu-id="806c8-171">X</span></span>    |     <span data-ttu-id="806c8-172">X</span><span class="sxs-lookup"><span data-stu-id="806c8-172">X</span></span>     |    <span data-ttu-id="806c8-173">X</span><span class="sxs-lookup"><span data-stu-id="806c8-173">X</span></span>    |
| <span data-ttu-id="806c8-174">[Rule (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-174">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="806c8-175">[Rule (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-175">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="806c8-176">X</span><span class="sxs-lookup"><span data-stu-id="806c8-176">X</span></span>    |
| <span data-ttu-id="806c8-177">[Requirements (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="806c8-177">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="806c8-178">X</span><span class="sxs-lookup"><span data-stu-id="806c8-178">X</span></span>    |
| <span data-ttu-id="806c8-179">[Set\*][]</span><span class="sxs-lookup"><span data-stu-id="806c8-179">[Set\*][]</span></span><br/><span data-ttu-id="806c8-180">[Sets (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="806c8-180">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="806c8-181">X</span><span class="sxs-lookup"><span data-stu-id="806c8-181">X</span></span>    |
| <span data-ttu-id="806c8-182">[Form\*][]</span><span class="sxs-lookup"><span data-stu-id="806c8-182">[Form\*][]</span></span><br/><span data-ttu-id="806c8-183">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="806c8-183">[formsettings\*][]</span></span>                                                              |         |           |    <span data-ttu-id="806c8-184">X</span><span class="sxs-lookup"><span data-stu-id="806c8-184">X</span></span>    |
| <span data-ttu-id="806c8-185">[Sets (Requirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="806c8-185">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="806c8-186">X</span><span class="sxs-lookup"><span data-stu-id="806c8-186">X</span></span>    |     <span data-ttu-id="806c8-187">X</span><span class="sxs-lookup"><span data-stu-id="806c8-187">X</span></span>     |         |
| <span data-ttu-id="806c8-188">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="806c8-188">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="806c8-189">X</span><span class="sxs-lookup"><span data-stu-id="806c8-189">X</span></span>    |     <span data-ttu-id="806c8-190">X</span><span class="sxs-lookup"><span data-stu-id="806c8-190">X</span></span>     |         |

<span data-ttu-id="806c8-191">_\*Ajouté dans le schéma de manifeste du complément Office version 1.1._</span><span class="sxs-lookup"><span data-stu-id="806c8-191">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<!-- Links for above table -->

[officeapp]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/officeapp?view=office-js
[id]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id
[version]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/version
[providername]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/providername
[defaultlocale]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultlocale
[displayname]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/displayname
[description]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/description
[iconurl]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/iconurl
[highresolutioniconurl]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/highresolutioniconurl
[defaultsettings (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultsettings
[sourcelocation (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation
[desktopsettings]: https://msdn.microsoft.com/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c(Office.15).aspx
[sourcelocation (mailapp)]: http://msdn.microsoft.com/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[permissions (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[permissions (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[permissions (mailapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[rule (rulecollection)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/rule
[rule (mailapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/rule
[requirements (mailapp)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements
[set*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/set
[sets (mailapprequirements)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sets
[form*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/form
[formsettings*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/formsettings
[sets (requirements)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sets
[hôtes\*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts
[hosts\*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts

## <a name="hosting-requirements"></a><span data-ttu-id="806c8-219">Configuration requise pour l’hébergement</span><span class="sxs-lookup"><span data-stu-id="806c8-219">Hosting requirements</span></span>

<span data-ttu-id="806c8-p102">Toutes les images des URI, telles que celles utilisées pour les [commandes du complément][], doivent prendre en charge la mise en cache. Le serveur qui héberge l’image ne doit pas renvoyer une `Cache-Control` en-tête spécifiant `no-cache`, `no-store`, ou des options similaires dans la réponse HTTP.</span><span class="sxs-lookup"><span data-stu-id="806c8-p102">All image URIs, such as those used for [add-in commands][], must support caching. The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="806c8-222">Toutes les URL, telles que les emplacements des fichiers source spécifiés dans l’élément [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation), doivent être **sécurisées par une protection SSL (HTTPS)**.</span><span class="sxs-lookup"><span data-stu-id="806c8-222">All URLs, such as the source file locations specified in the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="806c8-223">Bonnes pratiques pour l’envoi dans AppSource</span><span class="sxs-lookup"><span data-stu-id="806c8-223">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="806c8-p103">Vérifiez que l’ID du complément est un GUID valide et unique. Vous trouverez des outils de génération de GUID sur Internet pour vous aider à créer un GUID unique.</span><span class="sxs-lookup"><span data-stu-id="806c8-p103">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="806c8-p104">Compléments soumis au AppSource doivent également inclure l’élément [SupportUrl](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/supporturl) . Pour plus d’informations, voir [stratégies de Validation pour les compléments et les applications envoyées à AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span><span class="sxs-lookup"><span data-stu-id="806c8-p104">Add-ins submitted to AppSource must also include the [SupportUrl](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/supporturl) element. For more information, see [Validation policies for apps and add-ins submitted to AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span></span>

<span data-ttu-id="806c8-228">Utilisez uniquement l’élément [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) pour spécifier des domaines différents de celui spécifié dans l’élément [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) pour les scénarios d’authentification.</span><span class="sxs-lookup"><span data-stu-id="806c8-228">Only use the [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) element to specify domains other than the one specified in the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="806c8-229">Spécifier les domaines que vous souhaitez ouvrir dans la fenêtre de complément</span><span class="sxs-lookup"><span data-stu-id="806c8-229">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="806c8-p105">Toutefois, sur une plate-forme bureau, si votre complément tente d’accéder à une URL située dans un autre domaine que celui qui héberge la page initiale (comme indiqué dans l’élément [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) du fichier manifeste), cette URL s’ouvre dans une nouvelle fenêtre de navigateur en dehors du volet de complément de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="806c8-p105">When running in Office Online, your task pane can be navigated to any URL. However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office host application.</span></span>

<span data-ttu-id="806c8-p106">Pour remplacer ce comportement (bureau Office), spécifiez chaque domaine que vous voulez ouvrir dans la fenêtre Macro complémentaire dans la liste des domaines spécifiés dans l’élément [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) du fichier de manifeste. Si le complément tente d’accéder à une URL dans un domaine qui se trouve dans la liste, puis il s’ouvre dans le volet Office de bureau Office et Office Online. Si elle tente d’accéder à une URL qui n’est pas dans la liste, puis, dans Office de bureau, cette URL s’ouvre dans une nouvelle fenêtre de navigateur (à l’extérieur du volet complément).</span><span class="sxs-lookup"><span data-stu-id="806c8-p106">To override this (desktop Office) behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) element of the manifest file. If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both desktop Office and Office Online. If it tries to go to a URL that isn't in the list, then, in desktop Office, that URL opens in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="806c8-p107">Ce comportement s’applique uniquement au volet racine de la macro complémentaire. S’il existe un iframe incorporé dans la page complément, l’iframe peut être dirigée à une URL que si elle est répertoriée dans **AppDomains**, même dans Office de bureau.</span><span class="sxs-lookup"><span data-stu-id="806c8-p107">This behavior applies only to the root pane of the add-in. If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>

<span data-ttu-id="806c8-p108">L’exemple de manifeste XML suivant héberge sa page de complément principale dans le domaine `https://www.contoso.com` comme indiqué dans l’élément **SourceLocation**. Il indique également le domaine `https://www.northwindtraders.com` dans un élément [AppDomain](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomain) au sein de la liste d’éléments **AppDomains**. Si le complément ouvre une page dans le domaine www.northwindtraders.com, cette page s’ouvre dans le volet de complément.</span><span class="sxs-lookup"><span data-stu-id="806c8-p108">The following XML manifest example hosts its main add-in page in the `https://www.contoso.com` domain as specified in the **SourceLocation** element. It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomain) element within the **AppDomains** element list. If the add-in goes to a page in the www.northwindtraders.com domain, that page opens in the add-in pane, even in Office desktop.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="806c8-240">Exemples et schémas de fichier XML manifeste version 1.1</span><span class="sxs-lookup"><span data-stu-id="806c8-240">Manifest v1.1 XML file examples and schemas</span></span>
<span data-ttu-id="806c8-241">Les sections suivantes présentent des exemples de fichiers manifeste XML version 1.1 pour des compléments de contenu, de volet Office et Outlook.</span><span class="sxs-lookup"><span data-stu-id="806c8-241">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-panetabtabid-1"></a>[<span data-ttu-id="806c8-242">Volet Office</span><span class="sxs-lookup"><span data-stu-id="806c8-242">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="806c8-243">Schéma de manifeste d’application de volet Office</span><span class="sxs-lookup"><span data-stu-id="806c8-243">Task pane app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

<!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

<!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

<!--IMPORTANT! Id must be unique for your add-in. If you clone this manifest ensure that you change this id to your own GUID -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
   <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
   <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://i.imgur.com/oZFS95h.png" />

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
        <!-- Form factor. Currenly only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
            <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office Ribbon-->
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
                <!--Icons. Required sizes 16,31,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX -->
                <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFuncion or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  - <Action xsi:type="ExecuteFunction">
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
              </Icon>
              <Items>
                <Item id="Contoso.Menu.Item1">
                  <Label resid="Contoso.Item1.Label"/>
                  <Supertip>
                    <Title resid="Contoso.Item1.Label" />
                    <Description resid="Contoso.Item1.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
              <!-- If validating with XSD it needs to be at the end, we might change this before release -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon" DefaultValue="https://i.imgur.com/FkSShX9.png" />
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

# <a name="contenttabtabid-2"></a>[<span data-ttu-id="806c8-244">Contenu</span><span class="sxs-lookup"><span data-stu-id="806c8-244">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="806c8-245">Schéma de manifeste d’application de contenu</span><span class="sxs-lookup"><span data-stu-id="806c8-245">Content app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/ENUSIcon.png" />
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

# <a name="mailtabtabid-3"></a>[<span data-ttu-id="806c8-246">Messagerie</span><span class="sxs-lookup"><span data-stu-id="806c8-246">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="806c8-247">Schéma de manifeste d’application de messagerie</span><span class="sxs-lookup"><span data-stu-id="806c8-247">Mail app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">

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
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following line to specify    -->
  <!-- the web serverthat hosts the icon file. -->
  <IconUrl DefaultValue=
    "https://webserver/YouTube/YouTubeLogo.png"/>

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

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="806c8-248">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="806c8-248">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="806c8-p109">Pour résoudre les problèmes rencontrés avec votre manifeste, consultez la rubrique relative à la [validation et à la résolution des problèmes avec votre manifeste](../testing/troubleshoot-manifest.md). Vous apprendrez à valider le manifeste par rapport à la [définition de schéma XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) et à utiliser la journalisation runtime pour déboguer le manifeste.</span><span class="sxs-lookup"><span data-stu-id="806c8-p109">For troubleshooting issues with your manifest, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md). There, you will find information on how to validate the manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), and also how to use runtime logging to debug the manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="806c8-251">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="806c8-251">See also</span></span>

* <span data-ttu-id="806c8-252">[Création de commandes de complément dans votre manifeste][commandes de complément]</span><span class="sxs-lookup"><span data-stu-id="806c8-252">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="806c8-253">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="806c8-253">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="806c8-254">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="806c8-254">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="806c8-255">Référence de schéma pour les manifestes des compléments Office</span><span class="sxs-lookup"><span data-stu-id="806c8-255">Schema reference for Office Add-ins manifests</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [<span data-ttu-id="806c8-256">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="806c8-256">Validate and troubleshoot issues with your manifest</span></span>](../testing/troubleshoot-manifest.md)

[commandes de complément]: create-addin-commands.md
[add-in commands]: create-addin-commands.md