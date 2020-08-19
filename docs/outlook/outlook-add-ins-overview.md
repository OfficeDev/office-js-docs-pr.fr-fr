---
title: Présentation des compléments Outlook
description: Les compléments Outlook sont des intégrations conçues par des tiers dans Outlook à l’aide de notre plate-forme web.
ms.date: 08/18/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 83644823f4ca906f52cae430fa3a7f350dbf076c
ms.sourcegitcommit: e9f23a2857b90a7c17e3152292b548a13a90aa33
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/19/2020
ms.locfileid: "46803778"
---
# <a name="outlook-add-ins-overview"></a><span data-ttu-id="2be72-103">Présentation des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="2be72-103">Outlook add-ins overview</span></span>

<span data-ttu-id="2be72-104">Les compléments Outlook sont des intégrations conçues par des tiers dans Outlook à l’aide de notre plate-forme web.</span><span class="sxs-lookup"><span data-stu-id="2be72-104">Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform.</span></span> <span data-ttu-id="2be72-105">Les compléments Outlook comportent trois aspects clés :</span><span class="sxs-lookup"><span data-stu-id="2be72-105">Outlook add-ins have three key aspects:</span></span>

- <span data-ttu-id="2be72-106">La même logique complémentaire et commerciale fonctionne sur les ordinateurs de bureau (Outlook sur Windows et Mac), sur le web (Microsoft 365 et Outlook.com) et sur les téléphones portables.</span><span class="sxs-lookup"><span data-stu-id="2be72-106">The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Microsoft 365 and Outlook.com), and mobile.</span></span>
- <span data-ttu-id="2be72-107">Les compléments Outlook se composent d’un manifeste, qui décrit la manière dont le complément s’intègre dans Outlook (par exemple, un bouton ou un volet de tâches), ainsi que d’un code JavaScript/HTML, qui constitue l’interface utilisateur et la logique métier du complément.</span><span class="sxs-lookup"><span data-stu-id="2be72-107">Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.</span></span>
- <span data-ttu-id="2be72-108">Les compléments Outlook peuvent être acquis à partir d’[AppSource](https://appsource.microsoft.com) ou [chargés séparément](sideload-outlook-add-ins-for-testing.md) par les utilisateurs finals ou les administrateurs.</span><span class="sxs-lookup"><span data-stu-id="2be72-108">Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.</span></span>

<span data-ttu-id="2be72-109">Les compléments Outlook diffèrent des compléments COM ou VSTO, qui sont des intégrations plus anciennes spécifiques d’Outlook sous Windows.</span><span class="sxs-lookup"><span data-stu-id="2be72-109">Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows.</span></span> <span data-ttu-id="2be72-110">Contrairement aux compléments COM, les compléments Outlook ne comportent pas de code physiquement installé sur le périphérique de l’utilisateur ou du client Outlook.</span><span class="sxs-lookup"><span data-stu-id="2be72-110">Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client.</span></span> <span data-ttu-id="2be72-111">Pour un complément Outlook, Outlook lit le manifeste et raccorde les contrôles spécifiés dans l’interface utilisateur, puis charge le code JavaScript et HTML.</span><span class="sxs-lookup"><span data-stu-id="2be72-111">For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML.</span></span> <span data-ttu-id="2be72-112">Les composants web s’exécutent tous dans le contexte d’un navigateur dans un bac à sable (sandbox).</span><span class="sxs-lookup"><span data-stu-id="2be72-112">The web components all run in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="2be72-113">Les éléments Outlook qui prennent en charge les compléments incluent notamment les messages électroniques, les demandes de réunion, les réponses à des demandes de réunion, les annulations de réunion et les rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2be72-113">The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments.</span></span> <span data-ttu-id="2be72-114">Chaque complément Outlook définit le contexte dans lequel il est disponible, y compris les types d’éléments et si l’utilisateur lit ou compose un élément.</span><span class="sxs-lookup"><span data-stu-id="2be72-114">Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a><span data-ttu-id="2be72-115">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="2be72-115">Extension points</span></span>

<span data-ttu-id="2be72-p104">Les points d’extension correspondent à la manière dont les compléments sont intégrés à Outlook. Voici les méthodes possibles :</span><span class="sxs-lookup"><span data-stu-id="2be72-p104">Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:</span></span>

- <span data-ttu-id="2be72-p105">Les compléments peuvent indiquer des boutons qui apparaissent dans les surfaces de commande dans les messages et les rendez-vous. Pour plus d’informations, voir [Commandes de complément pour Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="2be72-p105">Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

    <span data-ttu-id="2be72-120">**Complément avec boutons de commande dans le ruban**</span><span class="sxs-lookup"><span data-stu-id="2be72-120">**An add-in with command buttons on the ribbon**</span></span>

    ![Forme sans interface utilisateur de commande de complément](../images/uiless-command-shape.png)

- <span data-ttu-id="2be72-p106">Les compléments peuvent désactiver les correspondances d’expressions régulières ou des entités détectées dans les messages et les rendez-vous. Pour plus d’informations, voir [Compléments Outlook contextuels](contextual-outlook-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="2be72-p106">Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).</span></span>

    <span data-ttu-id="2be72-124">**Complément contextuel pour une entité en surbrillance (adresse)**</span><span class="sxs-lookup"><span data-stu-id="2be72-124">**A contextual add-in for a highlighted entity (an address)**</span></span>

    ![Présente une application contextuelle dans une carte](../images/outlook-detected-entity-card.png)

> [!NOTE]
> <span data-ttu-id="2be72-126">[Les volets personnalisés sont déconseillés](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/). Veuillez donc vérifier que vous utilisez un point d’extension pris en charge.</span><span class="sxs-lookup"><span data-stu-id="2be72-126">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using a supported extension point.</span></span>

## <a name="mailbox-items-available-to-add-ins"></a><span data-ttu-id="2be72-127">Éléments de boîtes aux lettres disponibles pour les compléments</span><span class="sxs-lookup"><span data-stu-id="2be72-127">Mailbox items available to add-ins</span></span>

<span data-ttu-id="2be72-p107">Les compléments Outlook sont disponibles pour les messages ou les rendez-vous en mode de lecture ou de composition, mais pas pour d’autres types d’élément. Outlook ne les active pas si l’élément de message actuel, en mode de composition ou de lecture, fait partie des éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="2be72-p107">Outlook add-ins are available on messages or appointments while composing or reading, but not other item types. Outlook does not activate add-ins if the current message item, in a compose or read form, is one of the following:</span></span>

- <span data-ttu-id="2be72-p108">protégé par la Gestion des droits relatifs à l’information (IRM) ou chiffré par d’autres moyens de protection. Un message signé numériquement en est un exemple, puisque la signature numérique dépend de l’un de ces mécanismes ;</span><span class="sxs-lookup"><span data-stu-id="2be72-p108">Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.</span></span>

  > [!IMPORTANT]
  > - <span data-ttu-id="2be72-132">Les compléments s’activent sur les messages signés numériquement dans Outlook avec un abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="2be72-132">Add-ins activate on digitally signed messages in Outlook associated with a Microsoft 365 subscription.</span></span> <span data-ttu-id="2be72-133">Dans Windows, cette prise en charge a été introduite avec le build 8711.1000.</span><span class="sxs-lookup"><span data-stu-id="2be72-133">On Windows, this support was introduced with build 8711.1000.</span></span>
  >
  > - <span data-ttu-id="2be72-134">Démarrer avec Outlook build 13120.1000 sur Windows, les compléments peuvent désormais activer les éléments protégés par IRM.</span><span class="sxs-lookup"><span data-stu-id="2be72-134">Starting with Outlook build 13120.1000 on Windows, add-ins can now activate on items protected by IRM.</span></span> <span data-ttu-id="2be72-135">Pour plus d’informations sur cette fonctionnalité en mode aperçu, voir [Activation de complément sur les éléments protégés par la gestion des droits relatifs à l’information (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span><span class="sxs-lookup"><span data-stu-id="2be72-135">For more information about this feature in preview, see [Add-in activation on items protected by Information Rights Management (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span></span>

- <span data-ttu-id="2be72-136">un rapport ou une notification de remise qui a la classe de message IPM.Report.\* (notamment les rapports de remise et les notifications d’échec de remise, ainsi que les notifications de lecture, de non-lecture et de retard) ;</span><span class="sxs-lookup"><span data-stu-id="2be72-136">A delivery report or notification that has the message class IPM.Report.\*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.</span></span>

- <span data-ttu-id="2be72-137">un brouillon (aucun expéditeur n’y est affecté), ou dans le dossier Brouillons d’Outlook ;</span><span class="sxs-lookup"><span data-stu-id="2be72-137">A draft (does not have a sender assigned to it), or in the Outlook Drafts folder.</span></span>

- <span data-ttu-id="2be72-138">un fichier .msg ou .eml joint à un autre message ;</span><span class="sxs-lookup"><span data-stu-id="2be72-138">A .msg or .eml file which is an attachment to another message.</span></span>

- <span data-ttu-id="2be72-139">un fichier .msg ou .eml ouvert à partir du système de fichiers ;</span><span class="sxs-lookup"><span data-stu-id="2be72-139">A .msg or .eml file opened from the file system.</span></span>

- <span data-ttu-id="2be72-140">dans une boîte aux lettres partagée, dans la boîte aux lettres d’un autre utilisateur, dans une boîte aux lettres d’archivage ou dans un dossier public.</span><span class="sxs-lookup"><span data-stu-id="2be72-140">In a shared mailbox, in another user's mailbox, in an archive mailbox, or in a public folder.</span></span>

- <span data-ttu-id="2be72-141">utilise un formulaire personnalisé.</span><span class="sxs-lookup"><span data-stu-id="2be72-141">Using a custom form.</span></span>

<span data-ttu-id="2be72-142">En général, Outlook peut activer des compléments sous forme de lecture pour les éléments dans le dossier Éléments envoyés, à l'exception des compléments qui s’activent en fonction des correspondances de chaînes d'entités connues.</span><span class="sxs-lookup"><span data-stu-id="2be72-142">In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities.</span></span> <span data-ttu-id="2be72-143">Pour plus d'informations sur les raisons de ce problème, reportez-vous à la rubrique "Prise en charge pour les entités connues" dans [Faire correspondre des chaînes dans un élément Outlook en tant qu'entités connues](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="2be72-143">For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>

## <a name="supported-hosts"></a><span data-ttu-id="2be72-144">Hôtes pris en charge</span><span class="sxs-lookup"><span data-stu-id="2be72-144">Supported hosts</span></span>

<span data-ttu-id="2be72-145">Les add-ins Outlook sont pris en charge dans Outlook 2013 ou plus récent sur Windows, Outlook 2016 ou plus récent sur Mac, Outlook sur le web pour Exchange 2013 sur site et versions ultérieures, Outlook sur iOS, Outlook sur Android, et Outlook sur le web et Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="2be72-145">Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web and Outlook.com.</span></span> <span data-ttu-id="2be72-146">Les fonctionnalités les plus récentes ne sont pas toutes prises en charge dans tous les [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) à la fois.</span><span class="sxs-lookup"><span data-stu-id="2be72-146">Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time.</span></span> <span data-ttu-id="2be72-147">Reportez-vous aux articles et références API relatives à ces fonctionnalités pour savoir dans quels hôtes elles peuvent ou non être prises en charge.</span><span class="sxs-lookup"><span data-stu-id="2be72-147">Please refer to articles and API references for those features to see which hosts they may or may not be supported in.</span></span>


## <a name="get-started-building-outlook-add-ins"></a><span data-ttu-id="2be72-148">Commencer à créer des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="2be72-148">Get started building Outlook add-ins</span></span>

<span data-ttu-id="2be72-149">Pour commencer à créer des compléments Outlook, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="2be72-149">To get started building Outlook add-ins, try the following.</span></span>

- <span data-ttu-id="2be72-150">[Démarrage rapide](../quickstarts/outlook-quickstart.md) : créer un volet Office simple.</span><span class="sxs-lookup"><span data-stu-id="2be72-150">[Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.</span></span>
- <span data-ttu-id="2be72-151">[Didacticiel](../tutorials/outlook-tutorial.md) : découvrez comment créer un complément qui insère des gists GitHub dans un nouveau message.</span><span class="sxs-lookup"><span data-stu-id="2be72-151">[Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.</span></span>


## <a name="see-also"></a><span data-ttu-id="2be72-152">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2be72-152">See also</span></span>

- [<span data-ttu-id="2be72-153">Meilleures pratiques en matière de développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="2be72-153">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="2be72-154">Instructions de conception pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="2be72-154">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="2be72-155">Gérer les licences de compléments pour Office et SharePoint</span><span class="sxs-lookup"><span data-stu-id="2be72-155">License your Office and SharePoint Add-ins</span></span>](/office/dev/store/license-your-add-ins)
- [<span data-ttu-id="2be72-156">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="2be72-156">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="2be72-157">Mise à disposition de vos solutions sur AppSource et dans Office</span><span class="sxs-lookup"><span data-stu-id="2be72-157">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
