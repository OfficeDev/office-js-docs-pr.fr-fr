---
title: Fonctionnalité d’envoi des compléments Outlook
description: Permet de traiter un élément ou d’empêcher les utilisateurs d’effectuer certaines actions. Permet aussi aux compléments de définir certaines propriétés pendant l’envoi.
ms.date: 04/20/2021
localization_priority: Normal
ms.openlocfilehash: 126323527d74553aa7fd7e0c8cf1e5e5d89471ff
ms.sourcegitcommit: 691fa338029c9cbd9a7194d163f390c3321a0cd8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/23/2021
ms.locfileid: "51959179"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="7d941-103">Fonctionnalité d’envoi des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="7d941-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="7d941-p101">La fonctionnalité d’envoi des compléments Outlook vous permet de traiter un élément de message ou réunion, ou d’empêcher les utilisateurs d’effectuer certaines actions. Elle permet aussi aux compléments de définir certaines propriétés pendant l’envoi. Par exemple, vous pouvez utiliser la fonctionnalité d’envoi pour :</span><span class="sxs-lookup"><span data-stu-id="7d941-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="7d941-106">Empêcher un utilisateur d’envoyer des informations sensibles ou de laisser la ligne d’objet vide.</span><span class="sxs-lookup"><span data-stu-id="7d941-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="7d941-107">Ajouter un destinataire spécifique à la ligne CC dans les messages ou à la ligne destinataires facultatifs des réunions.</span><span class="sxs-lookup"><span data-stu-id="7d941-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

<span data-ttu-id="7d941-108">La fonctionnalité d’envoi est déclenchée par le type d’événement `ItemSend` et est sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7d941-108">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="7d941-109">Pour en savoir plus sur les limites de la fonctionnalité d’envoi, consultez la section [Limites](#limitations) plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="7d941-109">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="supported-clients-and-platforms"></a><span data-ttu-id="7d941-110">Clients et plateformes pris en charge</span><span class="sxs-lookup"><span data-stu-id="7d941-110">Supported clients and platforms</span></span>

<span data-ttu-id="7d941-111">Le tableau suivant présente les combinaisons client-serveur pris en charge pour la fonctionnalité d'envoi, y compris la mise à jour cumulative minimale requise, le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="7d941-111">The following table shows supported client-server combinations for the on-send feature, including the minimum required Cumulative Update where applicable.</span></span> <span data-ttu-id="7d941-112">Les combinaisons exclues ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="7d941-112">Excluded combinations are not supported.</span></span>

| <span data-ttu-id="7d941-113">Client</span><span class="sxs-lookup"><span data-stu-id="7d941-113">Client</span></span> | <span data-ttu-id="7d941-114">Exchange Online</span><span class="sxs-lookup"><span data-stu-id="7d941-114">Exchange Online</span></span> | <span data-ttu-id="7d941-115">Exchange 2016 local</span><span class="sxs-lookup"><span data-stu-id="7d941-115">Exchange 2016 on-premises</span></span><br><span data-ttu-id="7d941-116">(Mise à jour cumulative 6 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="7d941-116">(Cumulative Update 6 or later)</span></span> | <span data-ttu-id="7d941-117">Exchange 2019 local</span><span class="sxs-lookup"><span data-stu-id="7d941-117">Exchange 2019 on-premises</span></span><br><span data-ttu-id="7d941-118">(Mise à jour cumulative 1 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="7d941-118">(Cumulative Update 1 or later)</span></span> |
|---|:---:|:---:|:---:|
|<span data-ttu-id="7d941-119">Windows :</span><span class="sxs-lookup"><span data-stu-id="7d941-119">Windows:</span></span><br><span data-ttu-id="7d941-120">version 1910 (build 12130.20272) ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="7d941-120">version 1910 (build 12130.20272) or later</span></span>|<span data-ttu-id="7d941-121">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-121">Yes</span></span>|<span data-ttu-id="7d941-122">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-122">Yes</span></span>|<span data-ttu-id="7d941-123">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-123">Yes</span></span>|
|<span data-ttu-id="7d941-124">Mac :</span><span class="sxs-lookup"><span data-stu-id="7d941-124">Mac:</span></span><br><span data-ttu-id="7d941-125">build 16.30 à 16.46</span><span class="sxs-lookup"><span data-stu-id="7d941-125">build 16.30 to 16.46</span></span>|<span data-ttu-id="7d941-126">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-126">Yes</span></span>|<span data-ttu-id="7d941-127">Non</span><span class="sxs-lookup"><span data-stu-id="7d941-127">No</span></span>|<span data-ttu-id="7d941-128">Non</span><span class="sxs-lookup"><span data-stu-id="7d941-128">No</span></span>|
|<span data-ttu-id="7d941-129">Mac :</span><span class="sxs-lookup"><span data-stu-id="7d941-129">Mac:</span></span><br><span data-ttu-id="7d941-130">build 16.47 ou ultérieure</span><span class="sxs-lookup"><span data-stu-id="7d941-130">build 16.47 or later</span></span>|<span data-ttu-id="7d941-131">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-131">Yes</span></span>|<span data-ttu-id="7d941-132">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-132">Yes</span></span>|<span data-ttu-id="7d941-133">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-133">Yes</span></span>|
|<span data-ttu-id="7d941-134">Navigateur web :</span><span class="sxs-lookup"><span data-stu-id="7d941-134">Web browser:</span></span><br><span data-ttu-id="7d941-135">interface utilisateur Outlook moderne</span><span class="sxs-lookup"><span data-stu-id="7d941-135">modern Outlook UI</span></span>|<span data-ttu-id="7d941-136">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-136">Yes</span></span>|<span data-ttu-id="7d941-137">Non applicable</span><span class="sxs-lookup"><span data-stu-id="7d941-137">Not applicable</span></span>|<span data-ttu-id="7d941-138">Non applicable</span><span class="sxs-lookup"><span data-stu-id="7d941-138">Not applicable</span></span>|
|<span data-ttu-id="7d941-139">Navigateur web :</span><span class="sxs-lookup"><span data-stu-id="7d941-139">Web browser:</span></span><br><span data-ttu-id="7d941-140">interface utilisateur Outlook classique</span><span class="sxs-lookup"><span data-stu-id="7d941-140">classic Outlook UI</span></span>|<span data-ttu-id="7d941-141">Non applicable</span><span class="sxs-lookup"><span data-stu-id="7d941-141">Not applicable</span></span>|<span data-ttu-id="7d941-142">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-142">Yes</span></span>|<span data-ttu-id="7d941-143">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-143">Yes</span></span>|

> [!NOTE]
> <span data-ttu-id="7d941-144">La fonctionnalité d'envoi a été officiellement publiée dans l'ensemble de conditions requises 1.8 (pour plus d'informations, voir la prise en charge actuelle du serveur et du [client).](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)</span><span class="sxs-lookup"><span data-stu-id="7d941-144">The on-send feature was officially released in requirement set 1.8 (see [current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details).</span></span> <span data-ttu-id="7d941-145">Toutefois, notez que la matrice de prise en charge de la fonctionnalité est un sur-ensemble de l'ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="7d941-145">However, note that the feature's support matrix is a superset of the requirement set's.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7d941-146">Les applications qui utilisent la fonctionnalité d'envoi ne sont pas autorisées dans [AppSource.](https://appsource.microsoft.com)</span><span class="sxs-lookup"><span data-stu-id="7d941-146">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="7d941-147">Comment marche la fonctionnalité d’envoi ?</span><span class="sxs-lookup"><span data-stu-id="7d941-147">How does the on-send feature work?</span></span>

<span data-ttu-id="7d941-148">Vous pouvez utiliser la fonctionnalité d’envoi pour créer un complément Outlook qui intègre l’événement synchrone `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="7d941-148">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="7d941-149">Cet événement détecte le moment où l’utilisateur clique sur le bouton **Envoyer**(ou le bouton **Envoyer mise à jour** pour les réunions existantes) et peut servir à bloquer l’envoi de l’élément s’il n’est pas validé.</span><span class="sxs-lookup"><span data-stu-id="7d941-149">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="7d941-150">Par exemple, quand un utilisateur déclenche un événement d’envoi de message, un complément Outlook qui utilise la fonctionnalité d’envoi peut :</span><span class="sxs-lookup"><span data-stu-id="7d941-150">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="7d941-151">Lire et valider le contenu du message</span><span class="sxs-lookup"><span data-stu-id="7d941-151">Read and validate the email message contents</span></span>
- <span data-ttu-id="7d941-152">Vérifier que la ligne d’objet du message est remplie</span><span class="sxs-lookup"><span data-stu-id="7d941-152">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="7d941-153">Définir un destinataire prédéterminé</span><span class="sxs-lookup"><span data-stu-id="7d941-153">Set a predetermined recipient</span></span>

<span data-ttu-id="7d941-154">La validation est effectuée côté client dans Outlook lorsque l'événement d'envoi est déclenché et que le module dispose de 5 minutes avant son heure d'attente. Si la validation échoue, l'envoi de l'élément est bloqué et un message d'erreur s'affiche dans une barre d'informations qui invite l'utilisateur à prendre des mesures.</span><span class="sxs-lookup"><span data-stu-id="7d941-154">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

> [!NOTE]
> <span data-ttu-id="7d941-155">Dans Outlook sur le web, lorsque la fonctionnalité d'envoi est déclenchée dans un message en cours de composition dans l'onglet du navigateur Outlook, l'élément est publié dans sa propre fenêtre de navigateur ou onglet afin de terminer la validation et d'autres traitements.</span><span class="sxs-lookup"><span data-stu-id="7d941-155">In Outlook on the web, when the on-send feature is triggered in a message being composed within the Outlook browser tab, the item is popped out to its own browser window or tab in order to complete validation and other processing.</span></span>

<span data-ttu-id="7d941-156">La capture d’écran suivante montre une barre d’informations invitant l’expéditeur à renseigner l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="7d941-156">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![Capture d’écran qui montre un message d’erreur invitant l’utilisateur à renseigner l’objet du message](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="7d941-158">La capture d’écran suivante montre une barre d’informations informant l’expéditeur que des mots bloqués ont été trouvés.</span><span class="sxs-lookup"><span data-stu-id="7d941-158">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![Capture d’écran qui montre un message d’erreur indiquant que des mots bloqués ont été trouvés](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="7d941-160">Limites</span><span class="sxs-lookup"><span data-stu-id="7d941-160">Limitations</span></span>

<span data-ttu-id="7d941-161">Les limites de la fonctionnalité d’envoi sont les suivantes.</span><span class="sxs-lookup"><span data-stu-id="7d941-161">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="7d941-162">**Fonctionnalité d'envoi à l'envoi** &ndash; si vous appelez le [corps. AppendOnSendAsync dans](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) le handler d'envoi, une erreur est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="7d941-162">**Append-on-send** feature &ndash; If you call [body.AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) in the on-send handler, an error is returned.</span></span>
- <span data-ttu-id="7d941-163">**AppSource** &ndash; Vous ne pouvez pas publier de compléments Outlook qui utilisent la fonctionnalité d’envoi sur [AppSource](https://appsource.microsoft.com). car ils ne seront pas validés par AppSource.</span><span class="sxs-lookup"><span data-stu-id="7d941-163">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="7d941-164">Les compléments qui utilisent la fonctionnalité d’envoi doivent être déployés par les administrateurs.</span><span class="sxs-lookup"><span data-stu-id="7d941-164">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="7d941-165">**Manifeste** &ndash; Le complément prend en charge un seul événement `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="7d941-165">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="7d941-166">Si votre manifeste comprend plusieurs événements `ItemSend`, il ne sera pas validé.</span><span class="sxs-lookup"><span data-stu-id="7d941-166">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="7d941-p107">**Performances**&ndash; : plusieurs allers-retours vers le serveur web hébergeant le complément peuvent nuire aux performances du complément. Imaginez alors ce qu’occasionnerait la création de compléments nécessitant plusieurs opérations de messagerie ou réunions.</span><span class="sxs-lookup"><span data-stu-id="7d941-p107">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="7d941-169">**Envoyer plus tard** (Mac uniquement) &ndash; S’il y a des compléments d’envoi, la fonctionnalité **Envoyer plus tard** n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="7d941-169">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

<span data-ttu-id="7d941-170">En outre, il n'est pas recommandé d'appeler le handler d'événements d'envoi car la fermeture de l'élément doit se produire automatiquement une fois `item.close()` l'événement terminé.</span><span class="sxs-lookup"><span data-stu-id="7d941-170">Also, it's not recommended that you call `item.close()` in the on-send event handler as closing the item should happen automatically after the event is completed.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="7d941-171">Limites concernant le type ou le mode de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7d941-171">Mailbox type/mode limitations</span></span>

<span data-ttu-id="7d941-172">La fonctionnalité d’envoi est uniquement prise en charge pour les boîtes aux lettres utilisateur dans Outlook sur le web, sur Windows et sur Mac.</span><span class="sxs-lookup"><span data-stu-id="7d941-172">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="7d941-173">La fonctionnalité n’est pas prise en charge pour les types et modes de boîte aux lettres suivants.</span><span class="sxs-lookup"><span data-stu-id="7d941-173">The functionality is not currently supported for the following mailbox types and modes.</span></span>

- <span data-ttu-id="7d941-174">Boîtes aux lettres partagées\*</span><span class="sxs-lookup"><span data-stu-id="7d941-174">Shared mailboxes\*</span></span>
- <span data-ttu-id="7d941-175">Boîtes aux lettres de groupe</span><span class="sxs-lookup"><span data-stu-id="7d941-175">Group mailboxes</span></span>
- <span data-ttu-id="7d941-176">Mode hors connexion</span><span class="sxs-lookup"><span data-stu-id="7d941-176">Offline mode</span></span>

<span data-ttu-id="7d941-177">Outlook bloque l’envoi si la fonctionnalité d’envoi est activée pour ces scénarios de boîtes aux lettres.</span><span class="sxs-lookup"><span data-stu-id="7d941-177">Outlook won't allow sending if the on-send feature is enabled for these mailbox scenarios.</span></span> <span data-ttu-id="7d941-178">Toutefois, si un utilisateur répond à un e-mail dans une boîte aux lettres de groupe, le complément d’envoi n’est pas exécuté et le message est envoyé.</span><span class="sxs-lookup"><span data-stu-id="7d941-178">However, if a user responds to an email in a group mailbox, the on-send add-in won't run and the message will be sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7d941-179">\*La fonctionnalité d'envoi doit fonctionner sur des boîtes aux lettres ou des dossiers partagés si le add-in implémente également la prise en charge des scénarios d'accès [délégué.](delegate-access.md)</span><span class="sxs-lookup"><span data-stu-id="7d941-179">\* On-send functionality should work on shared mailboxes or folders if the add-in also [implements support for delegate access scenarios](delegate-access.md).</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="7d941-180">Compléments d’envoi multiples</span><span class="sxs-lookup"><span data-stu-id="7d941-180">Multiple on-send add-ins</span></span>

<span data-ttu-id="7d941-181">Si plusieurs compléments d’envoi sont installés, ils s’exécutent dans l’ordre dans lequel ils sont reçus par les API `getAppManifestCall` ou `getExtensibilityContext`.</span><span class="sxs-lookup"><span data-stu-id="7d941-181">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="7d941-182">Si le premier complément autorise l’envoi du message, le deuxième complément peut modifier un paramètre qui le bloque.</span><span class="sxs-lookup"><span data-stu-id="7d941-182">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="7d941-183">Par contre, le premier complément n’est pas réexécuté si les autres compléments installés autorisent l’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-183">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="7d941-184">Par exemple, Complément1 et Complément2 utilisent la fonctionnalité d’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-184">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="7d941-185">Complément1 est installé en premier, et Complément2 en deuxième.</span><span class="sxs-lookup"><span data-stu-id="7d941-185">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="7d941-186">Complément1 vérifie que le mot Fabrikam apparaît dans le message pour autoriser l’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-186">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="7d941-187">À l’inverse, Complément2 supprime toutes les occurrences du mot Fabrikam.</span><span class="sxs-lookup"><span data-stu-id="7d941-187">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="7d941-188">Le message est alors envoyé sans le mot Fabrikam (à cause de l’ordre d’installation de Complément1 et Complément2).</span><span class="sxs-lookup"><span data-stu-id="7d941-188">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="7d941-189">Déployer des compléments Outlook qui utilisent la fonctionnalité d’envoi</span><span class="sxs-lookup"><span data-stu-id="7d941-189">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="7d941-190">Nous recommandons aux administrateurs de déployer les compléments Outlook qui utilisent la fonctionnalité d’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-190">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="7d941-191">Les administrateurs doivent vérifier que le complément d’envoi :</span><span class="sxs-lookup"><span data-stu-id="7d941-191">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="7d941-192">est présent lors de l’ouverture d’un élément de composition (pour les e-mails : nouveau message, répondre ou transférer).</span><span class="sxs-lookup"><span data-stu-id="7d941-192">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="7d941-193">ne peut pas être fermé ou désactivé par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7d941-193">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="7d941-194">Installer des compléments Outlook qui utilisent la fonctionnalité d’envoi</span><span class="sxs-lookup"><span data-stu-id="7d941-194">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="7d941-195">Dans Outlook, la fonctionnalité d’envoi exige la configuration des compléments en fonction des types d’événement d’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-195">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="7d941-196">Sélectionnez la plateforme que vous voulez configurer.</span><span class="sxs-lookup"><span data-stu-id="7d941-196">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="7d941-197">Navigateur web – Outlook classique</span><span class="sxs-lookup"><span data-stu-id="7d941-197">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="7d941-198">Les compléments Outlook (classique) sur le web qui utilisent la fonctionnalité d’envoi s’exécutent pour les utilisateurs auxquels une stratégie de boîte aux lettres Outlook sur le web est attribuée, dont la valeur *OnSendAddinsEnabled* est définie sur **True**.</span><span class="sxs-lookup"><span data-stu-id="7d941-198">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="7d941-199">Pour installer un nouveau complément, exécutez les cmdlets Exchange Online PowerShell suivantes.</span><span class="sxs-lookup"><span data-stu-id="7d941-199">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="7d941-200">Pour découvrir comment utiliser PowerShell à distance afin de se connecter à Exchange Online, consultez la rubrique [Connexion à Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="7d941-200">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="7d941-201">Activer la fonctionnalité d’envoi</span><span class="sxs-lookup"><span data-stu-id="7d941-201">Enable the on-send feature</span></span>

<span data-ttu-id="7d941-202">Par défaut, la fonctionnalité d’envoi est désactivée.</span><span class="sxs-lookup"><span data-stu-id="7d941-202">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="7d941-203">Les administrateurs peuvent activer la fonctionnalité d’envoi en exécutant les cmdlets Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="7d941-203">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="7d941-204">Pour activer les compléments d’envoi pour tous les utilisateurs :</span><span class="sxs-lookup"><span data-stu-id="7d941-204">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="7d941-205">Créez une stratégie de boîte aux lettres Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="7d941-205">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="7d941-206">Les administrateurs peuvent utiliser une stratégie existante, mais la fonctionnalité d’envoi est uniquement prise en charge sur certains types de boîtes aux lettres.</span><span class="sxs-lookup"><span data-stu-id="7d941-206">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="7d941-207">La fonctionnalité d’envoi est bloquée par défaut sur les boîtes aux lettres non prises en charge dans Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="7d941-207">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="7d941-208">Activez la fonctionnalité d’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-208">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="7d941-209">Attribuez la stratégie à des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="7d941-209">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="7d941-210">Activer la fonctionnalité d’envoi pour un groupe d’utilisateurs</span><span class="sxs-lookup"><span data-stu-id="7d941-210">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="7d941-211">Pour activer la fonctionnalité d’envoi pour un groupe spécifique d’utilisateurs, suivez les étapes ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="7d941-211">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="7d941-212">Dans cet exemple, un administrateur souhaite uniquement activer un complément d’envoi Outlook sur le web dans un environnement réservé aux utilisateurs du service financier.</span><span class="sxs-lookup"><span data-stu-id="7d941-212">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="7d941-213">Créez une stratégie de boîte aux lettres Outlook sur le web pour le groupe.</span><span class="sxs-lookup"><span data-stu-id="7d941-213">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="7d941-214">Les administrateurs peuvent utiliser une stratégie existante, mais la fonctionnalité d’envoi est uniquement prise en charge sur certains types de boîtes aux lettres (pour en savoir plus, consultez la section [Limites concernant le type de boîte aux lettres](#multiple-on-send-add-ins) plus haut dans cet article).</span><span class="sxs-lookup"><span data-stu-id="7d941-214">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="7d941-215">La fonctionnalité d’envoi est bloquée par défaut sur les boîtes aux lettres non prises en charge dans Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="7d941-215">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="7d941-216">Activez la fonctionnalité d’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-216">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="7d941-217">Attribuez la stratégie à des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="7d941-217">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="7d941-218">vous devez attendre 60 minutes avant que la stratégie prenne effet. Sinon, redémarrez Internet Information Services (IIS).</span><span class="sxs-lookup"><span data-stu-id="7d941-218">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="7d941-219">Une fois la stratégie prise en compte, la fonctionnalité d’envoi est activée pour le groupe.</span><span class="sxs-lookup"><span data-stu-id="7d941-219">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="7d941-220">Désactiver la fonctionnalité d’envoi</span><span class="sxs-lookup"><span data-stu-id="7d941-220">Disable the on-send feature</span></span>

<span data-ttu-id="7d941-221">Pour désactiver la fonctionnalité d’envoi pour un utilisateur ou affecter une stratégie de boîte aux lettres Outlook sur le web dont l’indicateur est désactivé, exécutez les cmdlets suivantes.</span><span class="sxs-lookup"><span data-stu-id="7d941-221">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="7d941-222">Dans cet exemple, la stratégie de boîte aux lettres est *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="7d941-222">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="7d941-223">Pour en savoir plus sur l’utilisation de la cmdlet **Set-OwaMailboxPolicy** en vue de configurer des stratégies de boîte aux lettres Outlook sur le web existantes, consultez la rubrique [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="7d941-223">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="7d941-224">Pour désactiver la fonctionnalité d’envoi pour tous les utilisateurs auxquels une stratégie de boîte aux lettres Outlook sur le web spécifique est attribuée, exécutez les cmdlets suivantes.</span><span class="sxs-lookup"><span data-stu-id="7d941-224">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="7d941-225">Navigateur web – Outlook moderne</span><span class="sxs-lookup"><span data-stu-id="7d941-225">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="7d941-226">Les compléments pour Outlook sur le web (moderne) qui utilisent la fonctionnalité d’envoi doivent s’exécuter pour tous les utilisateurs qui les ont installés.</span><span class="sxs-lookup"><span data-stu-id="7d941-226">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="7d941-227">Toutefois, si les utilisateurs doivent exécuter des add-ins d'envoi pour répondre aux normes de conformité, la stratégie de boîte aux lettres doit avoir l'indicateur *OnSendAddinsEnabled* définie de sorte que la modification de l'élément n'est pas autorisée pendant le traitement des add-ins lors de `true` l'envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-227">However, if users are required to run on-send add-ins to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to `true` so that editing the item is not allowed while the add-ins are processing on send.</span></span>

<span data-ttu-id="7d941-228">Pour installer un nouveau complément, exécutez les cmdlets Exchange Online PowerShell suivantes.</span><span class="sxs-lookup"><span data-stu-id="7d941-228">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="7d941-229">Pour découvrir comment utiliser PowerShell à distance afin de se connecter à Exchange Online, consultez la rubrique [Connexion à Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="7d941-229">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-flag"></a><span data-ttu-id="7d941-230">Activer l'indicateur d'envoi</span><span class="sxs-lookup"><span data-stu-id="7d941-230">Enable the on-send flag</span></span>

<span data-ttu-id="7d941-231">Les administrateurs peuvent appliquer la conformité à l'envoi en exécutant des cmdlets Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="7d941-231">Administrators can enforce on-send compliance by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="7d941-232">Pour tous les utilisateurs, pour ne pas modifier pendant le traitement des add-ins d'envoi :</span><span class="sxs-lookup"><span data-stu-id="7d941-232">For all users, to disallow editing while on-send add-ins are processing:</span></span>

1. <span data-ttu-id="7d941-233">Créez une stratégie de boîte aux lettres Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="7d941-233">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="7d941-234">Les administrateurs peuvent utiliser une stratégie existante, mais la fonctionnalité d’envoi est uniquement prise en charge sur certains types de boîtes aux lettres.</span><span class="sxs-lookup"><span data-stu-id="7d941-234">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="7d941-235">La fonctionnalité d’envoi est bloquée par défaut sur les boîtes aux lettres non prises en charge dans Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="7d941-235">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="7d941-236">Appliquer la conformité lors de l'envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-236">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="7d941-237">Attribuez la stratégie à des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="7d941-237">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="turn-on-the-on-send-flag-for-a-group-of-users"></a><span data-ttu-id="7d941-238">Activer l'indicateur d'envoi pour un groupe d'utilisateurs</span><span class="sxs-lookup"><span data-stu-id="7d941-238">Turn on the on-send flag for a group of users</span></span>

<span data-ttu-id="7d941-239">Pour appliquer la conformité à l'envoi pour un groupe spécifique d'utilisateurs, les étapes sont les suivantes :</span><span class="sxs-lookup"><span data-stu-id="7d941-239">To enforce on-send compliance for a specific group of users, the steps are as follows.</span></span> <span data-ttu-id="7d941-240">Dans cet exemple, un administrateur souhaite uniquement activer une stratégie de complément d’envoi Outlook sur le web dans un environnement réservé aux utilisateurs du service financier.</span><span class="sxs-lookup"><span data-stu-id="7d941-240">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="7d941-241">Créez une stratégie de boîte aux lettres Outlook sur le web pour le groupe.</span><span class="sxs-lookup"><span data-stu-id="7d941-241">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="7d941-242">Les administrateurs peuvent utiliser une stratégie existante, mais la fonctionnalité d’envoi est uniquement prise en charge sur certains types de boîtes aux lettres (pour en savoir plus, consultez la section [Limites concernant le type de boîte aux lettres](#multiple-on-send-add-ins) plus haut dans cet article).</span><span class="sxs-lookup"><span data-stu-id="7d941-242">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="7d941-243">La fonctionnalité d’envoi est bloquée par défaut sur les boîtes aux lettres non prises en charge dans Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="7d941-243">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="7d941-244">Appliquer la conformité lors de l'envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-244">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="7d941-245">Attribuez la stratégie à des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="7d941-245">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="7d941-246">vous devez attendre 60 minutes avant que la stratégie prenne effet. Sinon, redémarrez Internet Information Services (IIS).</span><span class="sxs-lookup"><span data-stu-id="7d941-246">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="7d941-247">Lorsque la stratégie prend effet, la conformité à l'envoi est appliquée pour le groupe.</span><span class="sxs-lookup"><span data-stu-id="7d941-247">When the policy takes effect, on-send compliance will be enforced for the group.</span></span>

#### <a name="turn-off-the-on-send-flag"></a><span data-ttu-id="7d941-248">Désactiver l'indicateur d'envoi</span><span class="sxs-lookup"><span data-stu-id="7d941-248">Turn off the on-send flag</span></span>

<span data-ttu-id="7d941-249">Pour désactiver l'application de la conformité à l'envoi pour un utilisateur, affectez une stratégie de boîte aux lettres Outlook sur le web dont l'indicateur n'est pas activé en exécutant les cmdlets suivantes.</span><span class="sxs-lookup"><span data-stu-id="7d941-249">To turn off on-send compliance enforcement for a user, assign an Outlook on the web mailbox policy that does not have the flag enabled by running the following cmdlets.</span></span> <span data-ttu-id="7d941-250">Dans cet exemple, la stratégie de boîte aux lettres est *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="7d941-250">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="7d941-251">Pour en savoir plus sur l’utilisation de la cmdlet **Set-OwaMailboxPolicy** en vue de configurer des stratégies de boîte aux lettres Outlook sur le web existantes, consultez la rubrique [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="7d941-251">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="7d941-252">Pour désactiver l'application de la conformité à l'envoi pour tous les utilisateurs affectés à une stratégie de boîte aux lettres Outlook sur le web spécifique, exécutez les cmdlets suivantes.</span><span class="sxs-lookup"><span data-stu-id="7d941-252">To turn off on-send compliance enforcement for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[<span data-ttu-id="7d941-253">Windows</span><span class="sxs-lookup"><span data-stu-id="7d941-253">Windows</span></span>](#tab/windows)

<span data-ttu-id="7d941-254">Les compléments pour Outlook sur Windows qui utilisent la fonctionnalité d’envoi doivent s’exécuter pour tous les utilisateurs qui les ont installés.</span><span class="sxs-lookup"><span data-stu-id="7d941-254">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="7d941-255">Toutefois, si les utilisateurs sont obligés d’exécuter le complément pour respecter les normes de conformité, la stratégie de groupe **Désactiver l’envoi lorsque les extensions Web ne peuvent pas être chargées** doit être **Activée** sur chaque ordinateur concerné.</span><span class="sxs-lookup"><span data-stu-id="7d941-255">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="7d941-256">Pour définir des stratégies de boîte aux lettres, les administrateurs peuvent télécharger l'outil [Modèles](https://www.microsoft.com/download/details.aspx?id=49030) d'administration, puis accéder aux derniers modèles d'administration en exécutant l'Éditeur de stratégie de groupe local, **gpedit.msc**.</span><span class="sxs-lookup"><span data-stu-id="7d941-256">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy Editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="7d941-257">Rôle de la stratégie</span><span class="sxs-lookup"><span data-stu-id="7d941-257">What the policy does</span></span>

<span data-ttu-id="7d941-258">Pour des raisons de conformité, il se peut que les administrateurs doivent s’assurer que les utilisateurs ne peuvent pas envoyer de d’éléments message ou réunion tant que la dernière mise à jour du complément n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="7d941-258">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="7d941-259">Les administrateurs doivent activer la stratégie de groupe **Désactiver l’envoi lorsque les extensions Web ne peuvent pas être chargées**, de sorte que tous les compléments sont mis à jour à partir d’Exchange et disponibles pour vérifier que chaque élément message ou réunion respecte les règles et réglementations attendues lors de l’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-259">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="7d941-260">État de la stratégie</span><span class="sxs-lookup"><span data-stu-id="7d941-260">Policy status</span></span>|<span data-ttu-id="7d941-261">Résultat</span><span class="sxs-lookup"><span data-stu-id="7d941-261">Result</span></span>|
|---|---|
|<span data-ttu-id="7d941-262">Désactivé</span><span class="sxs-lookup"><span data-stu-id="7d941-262">Disabled</span></span>|<span data-ttu-id="7d941-263">Les manifestes actuellement téléchargés des applications d'envoi (pas nécessairement les versions les plus récentes) s'exécutent sur les éléments de message ou de réunion envoyés.</span><span class="sxs-lookup"><span data-stu-id="7d941-263">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="7d941-264">Il s'agit de l'état/comportement par défaut.</span><span class="sxs-lookup"><span data-stu-id="7d941-264">This is the default status/behavior.</span></span>|
|<span data-ttu-id="7d941-265">Activé</span><span class="sxs-lookup"><span data-stu-id="7d941-265">Enabled</span></span>|<span data-ttu-id="7d941-266">Une fois que les derniers manifestes des applications d'envoi sont téléchargés à partir d'Exchange, ils sont exécutés sur les éléments de message ou de réunion envoyés.</span><span class="sxs-lookup"><span data-stu-id="7d941-266">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="7d941-267">Sinon, l'envoi est bloqué.</span><span class="sxs-lookup"><span data-stu-id="7d941-267">Otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="7d941-268">Gérer la stratégie d’envoi</span><span class="sxs-lookup"><span data-stu-id="7d941-268">Manage the on-send policy</span></span>

<span data-ttu-id="7d941-269">Par défaut, la stratégie d’envoi est désactivée.</span><span class="sxs-lookup"><span data-stu-id="7d941-269">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="7d941-270">Les administrateurs peuvent activer la stratégie d’envoi en veillant à ce que le paramètre de la stratégie de groupe de l’utilisateur **Désactiver l'envoi lorsque les extensions Web ne sont pas chargées** soit **Activé**.</span><span class="sxs-lookup"><span data-stu-id="7d941-270">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="7d941-271">Pour désactiver la stratégie pour un utilisateur, l’administrateur doit la paramétrer sur **Désactivé**.</span><span class="sxs-lookup"><span data-stu-id="7d941-271">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="7d941-272">Pour gérer ce paramètre de stratégie, vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="7d941-272">To manage this policy setting, you can do the following:</span></span>

1. <span data-ttu-id="7d941-273">Téléchargez l’[outil de modèles d’administration](https://www.microsoft.com/download/details.aspx?id=49030).</span><span class="sxs-lookup"><span data-stu-id="7d941-273">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="7d941-274">Ouvrez l'Éditeur de stratégie de groupe local (**gpedit.msc**).</span><span class="sxs-lookup"><span data-stu-id="7d941-274">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="7d941-275">Accédez à **Configuration utilisateur > modèles d’administration > Microsoft Outlook 2016 > Sécurité > Centre de gestion de la confidentialité**.</span><span class="sxs-lookup"><span data-stu-id="7d941-275">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="7d941-276">Sélectionnez le paramètre **Désactiver l’envoi lorsque les extensions Web ne peuvent pas charger**.</span><span class="sxs-lookup"><span data-stu-id="7d941-276">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="7d941-277">Ouvrir le lien pour modifier le paramètre de stratégie.</span><span class="sxs-lookup"><span data-stu-id="7d941-277">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="7d941-278">Dans la fenêtre de dialogue **Désactiver l’envoi lorsque les extensions Web ne peuvent pas charger**, sélectionnez **Activée** ou **Désactivée**, puis sélectionnez **OK** ou **Appliquer** pour appliquer la mise à jour.</span><span class="sxs-lookup"><span data-stu-id="7d941-278">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="7d941-279">Mac</span><span class="sxs-lookup"><span data-stu-id="7d941-279">Mac</span></span>](#tab/unix)

<span data-ttu-id="7d941-280">Les compléments pour Outlook sur Mac qui utilisent la fonctionnalité d’envoi doivent s’exécuter pour tous les utilisateurs qui les ont installés.</span><span class="sxs-lookup"><span data-stu-id="7d941-280">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="7d941-281">Toutefois, si les utilisateurs sont obligés d’exécuter le complément pour respecter les normes de conformité, le paramètre de boîte aux lettres suivant doit être appliqué sur l’ordinateur de chaque utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7d941-281">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="7d941-282">Ce paramètre ou cette clé sont compatibles avec CFPreference, ce qui signifie qu’elle peut être définie à l’aide d’un logiciel de gestion d’entreprise pour Mac, tel que Jamf Pro.</span><span class="sxs-lookup"><span data-stu-id="7d941-282">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

||<span data-ttu-id="7d941-283">Valeur</span><span class="sxs-lookup"><span data-stu-id="7d941-283">Value</span></span>|
|:---|:---|
|<span data-ttu-id="7d941-284">**Domaine**</span><span class="sxs-lookup"><span data-stu-id="7d941-284">**Domain**</span></span>|<span data-ttu-id="7d941-285">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="7d941-285">com.microsoft.outlook</span></span>|
|<span data-ttu-id="7d941-286">**Clé**</span><span class="sxs-lookup"><span data-stu-id="7d941-286">**Key**</span></span>|<span data-ttu-id="7d941-287">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="7d941-287">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="7d941-288">**Type de données**</span><span class="sxs-lookup"><span data-stu-id="7d941-288">**DataType**</span></span>|<span data-ttu-id="7d941-289">Valeur booléenne</span><span class="sxs-lookup"><span data-stu-id="7d941-289">Boolean</span></span>|
|<span data-ttu-id="7d941-290">**Valeurs possibles**</span><span class="sxs-lookup"><span data-stu-id="7d941-290">**Possible values**</span></span>|<span data-ttu-id="7d941-291">false (par défaut)</span><span class="sxs-lookup"><span data-stu-id="7d941-291">false (default)</span></span><br><span data-ttu-id="7d941-292">true</span><span class="sxs-lookup"><span data-stu-id="7d941-292">true</span></span>|
|<span data-ttu-id="7d941-293">**Disponibilité**</span><span class="sxs-lookup"><span data-stu-id="7d941-293">**Availability**</span></span>|<span data-ttu-id="7d941-294">16.27</span><span class="sxs-lookup"><span data-stu-id="7d941-294">16.27</span></span>|
|<span data-ttu-id="7d941-295">**Commentaires**</span><span class="sxs-lookup"><span data-stu-id="7d941-295">**Comments**</span></span>|<span data-ttu-id="7d941-296">Cette clé crée une stratégie onSendMailbox.</span><span class="sxs-lookup"><span data-stu-id="7d941-296">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="7d941-297">Le rôle du paramètre</span><span class="sxs-lookup"><span data-stu-id="7d941-297">What the setting does</span></span>

<span data-ttu-id="7d941-298">Pour des raisons de conformité, il se peut que les administrateurs doivent s’assurer que les utilisateurs ne peuvent pas envoyer de d’éléments message ou réunion tant que la dernière mise à jour des compléments n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="7d941-298">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="7d941-299">Les administrateurs doivent activer la clé **OnSendAddinsWaitForLoad**, de sorte que tous les compléments sont mis à jour à partir d’Exchange et disponibles pour vérifier que chaque élément message ou réunion respecte les règles et réglementations attendues lors de l’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-299">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="7d941-300">État de la clé</span><span class="sxs-lookup"><span data-stu-id="7d941-300">Key's state</span></span>|<span data-ttu-id="7d941-301">Résultat</span><span class="sxs-lookup"><span data-stu-id="7d941-301">Result</span></span>|
|---|---|
|<span data-ttu-id="7d941-302">false</span><span class="sxs-lookup"><span data-stu-id="7d941-302">false</span></span>|<span data-ttu-id="7d941-303">Les manifestes actuellement téléchargés des applications d'envoi (pas nécessairement les versions les plus récentes) s'exécutent sur les éléments de message ou de réunion envoyés.</span><span class="sxs-lookup"><span data-stu-id="7d941-303">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="7d941-304">Il s'agit de l'état/comportement par défaut.</span><span class="sxs-lookup"><span data-stu-id="7d941-304">This is the default state/behavior.</span></span>|
|<span data-ttu-id="7d941-305">true</span><span class="sxs-lookup"><span data-stu-id="7d941-305">true</span></span>|<span data-ttu-id="7d941-306">Une fois que les derniers manifestes des applications d'envoi sont téléchargés à partir d'Exchange, ils sont exécutés sur les éléments de message ou de réunion envoyés.</span><span class="sxs-lookup"><span data-stu-id="7d941-306">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="7d941-307">Sinon, l'envoi est bloqué et le **bouton** Envoyer est désactivé.</span><span class="sxs-lookup"><span data-stu-id="7d941-307">Otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="7d941-308">Scénarios de la fonctionnalité d’envoi</span><span class="sxs-lookup"><span data-stu-id="7d941-308">On-send feature scenarios</span></span>

<span data-ttu-id="7d941-309">Voici tous les scénarios pris en charge et non pour les compléments qui utilisent la fonctionnalité d’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-309">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="7d941-310">La fonctionnalité d’envoi est activée sur la boîte aux lettres de l’utilisateur, mais aucun complément n’est installé.</span><span class="sxs-lookup"><span data-stu-id="7d941-310">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="7d941-311">Dans ce scénario, l’utilisateur peut envoyer des éléments message ou réunion sans l’exécution des compléments.</span><span class="sxs-lookup"><span data-stu-id="7d941-311">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="7d941-312">La fonctionnalité d’envoi est activée sur la boîte aux lettres de l’utilisateur et les compléments qui prennent en charge cette fonctionnalité sont installés et activés</span><span class="sxs-lookup"><span data-stu-id="7d941-312">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="7d941-313">Les compléments s’exécutent pendant l’événement d’envoi pour autoriser ou empêcher l’utilisateur d’envoyer son message.</span><span class="sxs-lookup"><span data-stu-id="7d941-313">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="7d941-314">Délégation de boîte aux lettres, où la Boîte aux lettres 1 dispose des autorisations d’accès total à la Boîte aux lettres 2</span><span class="sxs-lookup"><span data-stu-id="7d941-314">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="7d941-315">Navigateur web (Outlook classique)</span><span class="sxs-lookup"><span data-stu-id="7d941-315">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="7d941-316">Scénario</span><span class="sxs-lookup"><span data-stu-id="7d941-316">Scenario</span></span>|<span data-ttu-id="7d941-317">Fonctionnalité d’envoi (Boîte aux lettres 1)</span><span class="sxs-lookup"><span data-stu-id="7d941-317">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="7d941-318">Fonctionnalité d’envoi (Boîte aux lettres 2)</span><span class="sxs-lookup"><span data-stu-id="7d941-318">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="7d941-319">Session web Outlook (classique)</span><span class="sxs-lookup"><span data-stu-id="7d941-319">Outlook web session (classic)</span></span>|<span data-ttu-id="7d941-320">Résultat</span><span class="sxs-lookup"><span data-stu-id="7d941-320">Result</span></span>|<span data-ttu-id="7d941-321">Pris en charge ?</span><span class="sxs-lookup"><span data-stu-id="7d941-321">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="7d941-322">1</span><span class="sxs-lookup"><span data-stu-id="7d941-322">1</span></span>|<span data-ttu-id="7d941-323">Activé</span><span class="sxs-lookup"><span data-stu-id="7d941-323">Enabled</span></span>|<span data-ttu-id="7d941-324">Activé</span><span class="sxs-lookup"><span data-stu-id="7d941-324">Enabled</span></span>|<span data-ttu-id="7d941-325">Nouvelle session</span><span class="sxs-lookup"><span data-stu-id="7d941-325">New session</span></span>|<span data-ttu-id="7d941-326">La boîte aux lettres 1 ne peut pas envoyer un message ou un élément de réunion provenant de la boîte aux lettres 2.</span><span class="sxs-lookup"><span data-stu-id="7d941-326">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="7d941-p135">N’est pas pris en charge actuellement. Pour y remédier, utilisez le scénario 3.</span><span class="sxs-lookup"><span data-stu-id="7d941-p135">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="7d941-329">2</span><span class="sxs-lookup"><span data-stu-id="7d941-329">2</span></span>|<span data-ttu-id="7d941-330">Désactivé</span><span class="sxs-lookup"><span data-stu-id="7d941-330">Disabled</span></span>|<span data-ttu-id="7d941-331">Activé</span><span class="sxs-lookup"><span data-stu-id="7d941-331">Enabled</span></span>|<span data-ttu-id="7d941-332">Nouvelle session</span><span class="sxs-lookup"><span data-stu-id="7d941-332">New session</span></span>|<span data-ttu-id="7d941-333">La boîte aux lettres 1 ne peut pas envoyer un message ou un élément de réunion provenant de la boîte aux lettres 2.</span><span class="sxs-lookup"><span data-stu-id="7d941-333">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="7d941-p136">N’est pas pris en charge actuellement. Pour y remédier, utilisez le scénario 3.</span><span class="sxs-lookup"><span data-stu-id="7d941-p136">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="7d941-336">3</span><span class="sxs-lookup"><span data-stu-id="7d941-336">3</span></span>|<span data-ttu-id="7d941-337">Activé</span><span class="sxs-lookup"><span data-stu-id="7d941-337">Enabled</span></span>|<span data-ttu-id="7d941-338">Activé</span><span class="sxs-lookup"><span data-stu-id="7d941-338">Enabled</span></span>|<span data-ttu-id="7d941-339">Même session</span><span class="sxs-lookup"><span data-stu-id="7d941-339">Same session</span></span>|<span data-ttu-id="7d941-340">Les compléments d’envoi attribués à la boîte aux lettres 1 exécutent la fonctionnalité d’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-340">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="7d941-341">Pris en charge.</span><span class="sxs-lookup"><span data-stu-id="7d941-341">Supported.</span></span>|
|<span data-ttu-id="7d941-342">4 </span><span class="sxs-lookup"><span data-stu-id="7d941-342">4</span></span>|<span data-ttu-id="7d941-343">Activé</span><span class="sxs-lookup"><span data-stu-id="7d941-343">Enabled</span></span>|<span data-ttu-id="7d941-344">Désactivé</span><span class="sxs-lookup"><span data-stu-id="7d941-344">Disabled</span></span>|<span data-ttu-id="7d941-345">Nouvelle session</span><span class="sxs-lookup"><span data-stu-id="7d941-345">New session</span></span>|<span data-ttu-id="7d941-346">Aucun complément d’envoi ne s’exécute ; un message ou un élément de réunion est envoyé.</span><span class="sxs-lookup"><span data-stu-id="7d941-346">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="7d941-347">Pris en charge.</span><span class="sxs-lookup"><span data-stu-id="7d941-347">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="7d941-348">Navigateur web (Outlook moderne), Windows, Mac</span><span class="sxs-lookup"><span data-stu-id="7d941-348">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="7d941-349">Pour appliquer l’envoi, les administrateurs doivent s’assurer que la stratégie a été activée sur les deux boîtes aux lettres.</span><span class="sxs-lookup"><span data-stu-id="7d941-349">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="7d941-350">Pour plus d’informations sur la prise en charge de l’accès délégué dans un complément, voir [Activer les scénarios d’accès délégué dans un complément Outlook](delegate-access.md).</span><span class="sxs-lookup"><span data-stu-id="7d941-350">To learn how to support delegate access in an add-in, see [Enable delegate access scenarios in an Outlook add-in](delegate-access.md).</span></span>

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a><span data-ttu-id="7d941-351">Le Groupe 1 est une boîte aux lettres de groupe moderne et la Boîte aux lettres d’utilisateur 1 est membre du Groupe 1</span><span class="sxs-lookup"><span data-stu-id="7d941-351">Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1</span></span>

<br/>

|<span data-ttu-id="7d941-352">Scénario</span><span class="sxs-lookup"><span data-stu-id="7d941-352">Scenario</span></span>|<span data-ttu-id="7d941-353">Stratégie d’envoi de la boîte aux lettres 1</span><span class="sxs-lookup"><span data-stu-id="7d941-353">Mailbox 1 on-send policy</span></span>|<span data-ttu-id="7d941-354">Compléments d’envoi activés ?</span><span class="sxs-lookup"><span data-stu-id="7d941-354">On-send add-ins enabled?</span></span>|<span data-ttu-id="7d941-355">Action de la boîte aux lettres 1</span><span class="sxs-lookup"><span data-stu-id="7d941-355">Mailbox 1 action</span></span>|<span data-ttu-id="7d941-356">Résultat</span><span class="sxs-lookup"><span data-stu-id="7d941-356">Result</span></span>|<span data-ttu-id="7d941-357">Pris en charge ?</span><span class="sxs-lookup"><span data-stu-id="7d941-357">Supported?</span></span>|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|<span data-ttu-id="7d941-358">1</span><span class="sxs-lookup"><span data-stu-id="7d941-358">1</span></span>|<span data-ttu-id="7d941-359">Activé</span><span class="sxs-lookup"><span data-stu-id="7d941-359">Enabled</span></span>|<span data-ttu-id="7d941-360">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-360">Yes</span></span>|<span data-ttu-id="7d941-361">La Boîte aux lettres 1 compose un nouveau message ou réunion pour le Groupe 1.</span><span class="sxs-lookup"><span data-stu-id="7d941-361">Mailbox 1 composes new message or meeting to Group 1.</span></span>|<span data-ttu-id="7d941-362">Les compléments d’envoi s’exécutent pendant l’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-362">On-send add-ins run during send.</span></span>|<span data-ttu-id="7d941-363">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-363">Yes</span></span>|
|<span data-ttu-id="7d941-364">2</span><span class="sxs-lookup"><span data-stu-id="7d941-364">2</span></span>|<span data-ttu-id="7d941-365">Activé</span><span class="sxs-lookup"><span data-stu-id="7d941-365">Enabled</span></span>|<span data-ttu-id="7d941-366">Oui</span><span class="sxs-lookup"><span data-stu-id="7d941-366">Yes</span></span>|<span data-ttu-id="7d941-367">La boîte aux lettres 1 compose un nouveau message ou réunion pour le Groupe 1, dans la fenêtre du Groupe 1 dans Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="7d941-367">Mailbox 1 composes a new message or meeting to Group 1 within Group 1's group window in Outlook on the web.</span></span>|<span data-ttu-id="7d941-368">Les compléments d’envoi ne s’exécutent pas pendant l’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-368">On-send add-ins do not run during send.</span></span>|<span data-ttu-id="7d941-369">N’est pas pris en charge actuellement.</span><span class="sxs-lookup"><span data-stu-id="7d941-369">Not currently supported.</span></span> <span data-ttu-id="7d941-370">Pour y remédier, utilisez le scénario 1.</span><span class="sxs-lookup"><span data-stu-id="7d941-370">As a workaround, use scenario 1.</span></span>|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="7d941-371">La fonctionnalité/stratégie d’envoi est activée sur la boîte aux lettres de l’utilisateur, les compléments qui prennent en charge cette fonctionnalité sont installés et activés et le mode hors connexion est activé</span><span class="sxs-lookup"><span data-stu-id="7d941-371">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="7d941-372">Les compléments d’envoi s’exécutent en fonction de l’état en ligne de l’utilisateur, du serveur principal du complément et d’Exchange.</span><span class="sxs-lookup"><span data-stu-id="7d941-372">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="7d941-373">État de l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="7d941-373">User's state</span></span>

<span data-ttu-id="7d941-374">Les compléments d’envoi s’exécutent pendant l’envoi, si l’utilisateur est en ligne.</span><span class="sxs-lookup"><span data-stu-id="7d941-374">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="7d941-375">Si l’utilisateur est hors ligne, les compléments d’envoi ne s’exécutent pas pendant l’envoi et l’élément message ou réunion n’est pas envoyé.</span><span class="sxs-lookup"><span data-stu-id="7d941-375">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="7d941-376">État du serveur de complément</span><span class="sxs-lookup"><span data-stu-id="7d941-376">Add-in backend's state</span></span>

<span data-ttu-id="7d941-377">Un complément sur envoi s’exécute si son serveur principal est en ligne et joignable.</span><span class="sxs-lookup"><span data-stu-id="7d941-377">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="7d941-378">Si le serveur principal est hors connexion, l’envoi est désactivé.</span><span class="sxs-lookup"><span data-stu-id="7d941-378">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="7d941-379">État d’Exchange</span><span class="sxs-lookup"><span data-stu-id="7d941-379">Exchange's state</span></span>

<span data-ttu-id="7d941-380">Les compléments d’envoi s’exécutent pendant l’envoi, si le serveur Exchange est en ligne et joignable.</span><span class="sxs-lookup"><span data-stu-id="7d941-380">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="7d941-381">Si le complément sur envoi ne peut pas accéder à Exchange et que la stratégie ou l’applet de commande applicable sont activés, l’envoi est désactivé.</span><span class="sxs-lookup"><span data-stu-id="7d941-381">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="7d941-382">Sur Mac en mode hors connexion, le bouton **Envoyer** (ou le bouton **Envoyer mise à jour** pour les réunions existantes) est désactivé et une notification indique que l’organisation n’autorise pas l’envoi lorsque l’utilisateur est hors connexion.</span><span class="sxs-lookup"><span data-stu-id="7d941-382">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>

### <a name="user-can-edit-item-while-on-send-add-ins-are-working-on-it"></a><span data-ttu-id="7d941-383">L'utilisateur peut modifier l'élément pendant que les modules d'envoi y travaillent</span><span class="sxs-lookup"><span data-stu-id="7d941-383">User can edit item while on-send add-ins are working on it</span></span>

<span data-ttu-id="7d941-384">Pendant que les modules d'envoi traitent un élément, l'utilisateur peut modifier l'élément en ajoutant, par exemple, du texte inapproprié ou des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7d941-384">While on-send add-ins are processing an item, the user can edit the item by adding, for example, inappropriate text or attachments.</span></span> <span data-ttu-id="7d941-385">Si vous souhaitez empêcher l'utilisateur de modifier l'élément pendant que votre application est en cours de traitement lors de l'envoi, vous pouvez implémenter une solution de contournement à l'aide d'une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="7d941-385">If you want to prevent the user from editing the item while your add-in is processing on send, you can implement a workaround using a dialog.</span></span> <span data-ttu-id="7d941-386">Cette solution de contournement peut être utilisée dans Outlook sur le web (classique), Windows et Mac.</span><span class="sxs-lookup"><span data-stu-id="7d941-386">This workaround can be used in Outlook on the web (classic), Windows, and Mac.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7d941-387">Outlook sur le web moderne : pour empêcher l'utilisateur de modifier l'élément pendant que votre application est en cours de traitement lors de l'envoi, vous devez définir l'indicateur *OnSendAddinsEnabled* sur tel que décrit dans la section Installer les applications Outlook qui utilisent la section d'envoi plus tôt dans cet `true` article. [](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send)</span><span class="sxs-lookup"><span data-stu-id="7d941-387">Modern Outlook on the web: To prevent the user from editing the item while your add-in is processing on send, you should set the *OnSendAddinsEnabled* flag to `true` as described in the [Install Outlook add-ins that use on-send](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send) section earlier in this article.</span></span>

<span data-ttu-id="7d941-388">Dans votre handler d'envoi :</span><span class="sxs-lookup"><span data-stu-id="7d941-388">In your on-send handler:</span></span>

1. <span data-ttu-id="7d941-389">Appelez [displayDialogAsync pour ouvrir](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) une boîte de dialogue afin que les clics de souris et les frappes soient désactivés.</span><span class="sxs-lookup"><span data-stu-id="7d941-389">Call [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) to open a dialog so that mouse clicks and keystrokes are disabled.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="7d941-390">Pour obtenir ce comportement dans Outlook sur le web classique, vous devez définir la propriété [displayInIframe](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) dans le paramètre `true` `options` de `displayDialogAsync` l'appel.</span><span class="sxs-lookup"><span data-stu-id="7d941-390">To get this behavior in classic Outlook on the web, you should set the [displayInIframe property](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) to `true` in the `options` parameter of the `displayDialogAsync` call.</span></span>

1. <span data-ttu-id="7d941-391">Implémenter le traitement de l'élément.</span><span class="sxs-lookup"><span data-stu-id="7d941-391">Implement processing of the item.</span></span>
1. <span data-ttu-id="7d941-392">Fermez la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="7d941-392">Close the dialog.</span></span> <span data-ttu-id="7d941-393">En outre, traitez ce qui se produit si l'utilisateur ferme la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="7d941-393">Also, handle what happens if the user closes the dialog.</span></span>

## <a name="code-examples"></a><span data-ttu-id="7d941-394">Exemples de code</span><span class="sxs-lookup"><span data-stu-id="7d941-394">Code examples</span></span>

<span data-ttu-id="7d941-395">Les exemples de code ci-dessous vous montrent comment créer un complément d’envoi simple.</span><span class="sxs-lookup"><span data-stu-id="7d941-395">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="7d941-396">Pour télécharger l’exemple de code sur lequel se basent ces exemples, consultez l’article [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span><span class="sxs-lookup"><span data-stu-id="7d941-396">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

> [!TIP]
> <span data-ttu-id="7d941-397">Si vous utilisez une boîte de dialogue avec l'événement d'envoi, veillez à fermer la boîte de dialogue avant de terminer l'événement.</span><span class="sxs-lookup"><span data-stu-id="7d941-397">If you use a dialog with the on-send event, make sure to close the dialog before completing the event.</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="7d941-398">Manifeste, remplacement de version et événement</span><span class="sxs-lookup"><span data-stu-id="7d941-398">Manifest, version override, and event</span></span>

<span data-ttu-id="7d941-399">L’exemple de code [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) comprend deux manifestes :</span><span class="sxs-lookup"><span data-stu-id="7d941-399">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="7d941-400">`Contoso Message Body Checker.xml` &ndash; : montre comment vérifier la présence de mots non autorisés ou d’informations sensibles dans le corps d’un message pendant l’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-400">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="7d941-401">`Contoso Subject and CC Checker.xml` &ndash; : montre comment ajouter un destinataire à la ligne Cc et vérifier que le message comporte une ligne d’objet pendant l’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-401">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="7d941-402">Dans le fichier manifeste `Contoso Message Body Checker.xml`, insérez le fichier de fonction et le nom de la fonction qui doit être appelée lors d’un événement `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="7d941-402">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="7d941-403">L’opération s’exécute de façon synchrone.</span><span class="sxs-lookup"><span data-stu-id="7d941-403">The operation runs synchronously.</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case, the function validateBody will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateBody" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

> [!IMPORTANT]
> <span data-ttu-id="7d941-404">Si vous utilisez Visual Studio 2019 pour développer votre add-in d'envoi, vous pouvez obtenir un avertissement de validation comme suit : « Il s'agit d'un xsi:type ' non valide http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events ». Pour contourner ce besoin, vous aurez besoin d'une version plus récente de MailAppVersionOverridesV1_1.xsd qui a été fournie en tant que gist GitHub dans un blog à propos de cet [avertissement](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span><span class="sxs-lookup"><span data-stu-id="7d941-404">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="7d941-405">Pour le fichier manifeste `Contoso Subject and CC Checker.xml`, l’exemple suivant montre le fichier de fonction et le nom de la fonction à appeler dans l’événement d’envoi du message.</span><span class="sxs-lookup"><span data-stu-id="7d941-405">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case the function validateSubjectAndCC will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateSubjectAndCC" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

<br/>

<span data-ttu-id="7d941-406">L’API d’envoi nécessite `VersionOverrides v1_1`.</span><span class="sxs-lookup"><span data-stu-id="7d941-406">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="7d941-407">L’exemple vous montre comment ajouter le nœud `VersionOverrides` dans votre manifeste.</span><span class="sxs-lookup"><span data-stu-id="7d941-407">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On-send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="7d941-408">Pour plus d’informations, voir les commandes suivantes :</span><span class="sxs-lookup"><span data-stu-id="7d941-408">For more information, see the following:</span></span>
> - [<span data-ttu-id="7d941-409">Manifestes de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="7d941-409">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="7d941-410">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="7d941-410">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="7d941-411">Les objets `Event` et `item` et les méthodes `body.getAsync` et `body.setAsync`</span><span class="sxs-lookup"><span data-stu-id="7d941-411">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="7d941-412">Pour accéder au message ou élément de réunion sélectionné (dans cet exemple, le message que vous venez de composer), utilisez l’espace de noms `Office.context.mailbox.item`.</span><span class="sxs-lookup"><span data-stu-id="7d941-412">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="7d941-413">L’événement `ItemSend` est automatiquement transmis via la fonctionnalité d’envoi vers la fonction spécifiée dans le manifeste &mdash;,dans cet exemple, la fonction `validateBody`.</span><span class="sxs-lookup"><span data-stu-id="7d941-413">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

```js
var mailboxItem;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateBody(event) {
    mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
}
```

<span data-ttu-id="7d941-414">Le corps actuel de la fonction `validateBody` s’affiche dans le format spécifié (HTML) et transmet l’objet « event » `ItemSend` auquel le code souhaite accéder avec la méthode du rappel.</span><span class="sxs-lookup"><span data-stu-id="7d941-414">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="7d941-415">En plus de la méthode `getAsync`, l’objet `Body` fournit également une méthode `setAsync` utile pour remplacer le corps du message par le texte spécifié.</span><span class="sxs-lookup"><span data-stu-id="7d941-415">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="7d941-416">Pour en savoir plus, consultez les articles relatifs à l’objet [Event](/javascript/api/office/office.addincommands.event) et à la méthode [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="7d941-416">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="7d941-417">Objet `NotificationMessages` et méthode `event.completed`</span><span class="sxs-lookup"><span data-stu-id="7d941-417">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="7d941-418">La fonction `checkBodyOnlyOnSendCallBack` utilise une expression régulière pour déterminer si le corps du message contient des mots bloqués.</span><span class="sxs-lookup"><span data-stu-id="7d941-418">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="7d941-419">Si elle trouve une correspondance dans un tableau de mots bloqués, il bloque l’envoi du message et avertit l’expéditeur via la barre d’informations.</span><span class="sxs-lookup"><span data-stu-id="7d941-419">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="7d941-420">Pour ce faire, il utilise la propriété `notificationMessages` de l'objet `Item` pour renvoyer un objet `NotificationMessages`.</span><span class="sxs-lookup"><span data-stu-id="7d941-420">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="7d941-421">Il ajoute ensuite une notification à l’élément en appelant la méthode `addAsync`, comme illustré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="7d941-421">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

```js
// Determine whether the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allow sending.
// <param name="asyncResult">ItemSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
    var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
    var wordExpression = listOfBlockedWords.join('|');

    // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
    // i to perform case-insensitive search.
    var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
    var checkBody = regexCheck.test(asyncResult.value);

    if (checkBody) {
        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
        // Block send.
        asyncResult.asyncContext.completed({ allowEvent: false });
    }

    // Allow send.
    asyncResult.asyncContext.completed({ allowEvent: true });
}
```

<span data-ttu-id="7d941-422">Voici les paramètres pour la méthode `addAsync` :</span><span class="sxs-lookup"><span data-stu-id="7d941-422">The following are the parameters for the `addAsync` method:</span></span>

- <span data-ttu-id="7d941-423">`NoSend` &ndash; : chaîne correspondant à une clé spécifiée par un développeur pour référencer un message de notification.</span><span class="sxs-lookup"><span data-stu-id="7d941-423">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="7d941-424">Vous pouvez l’utiliser pour modifier ce message ultérieurement.</span><span class="sxs-lookup"><span data-stu-id="7d941-424">You can use it to modify this message later.</span></span> <span data-ttu-id="7d941-425">La clé ne peut pas avoir plus de 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="7d941-425">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="7d941-426">`type`&ndash; : l’une des propriétés du paramètre d’objet JSON.</span><span class="sxs-lookup"><span data-stu-id="7d941-426">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="7d941-427">Représente le type d’un message ; les types correspondent aux valeurs de l’énumération [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype).</span><span class="sxs-lookup"><span data-stu-id="7d941-427">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="7d941-428">Les valeurs possibles sont Indicateur de progression, Message d’information ou Message d’erreur.</span><span class="sxs-lookup"><span data-stu-id="7d941-428">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="7d941-429">Dans cet exemple, `type` est un message d’erreur.</span><span class="sxs-lookup"><span data-stu-id="7d941-429">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="7d941-430">`message`&ndash; : l’une des propriétés du paramètre d’objet JSON.</span><span class="sxs-lookup"><span data-stu-id="7d941-430">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="7d941-431">Dans cet exemple, `message` correspond au texte du message de notification.</span><span class="sxs-lookup"><span data-stu-id="7d941-431">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="7d941-432">Pour signaler que le complément a terminé le traitement de l’événement `ItemSend` déclenché par l’opération d’envoi, appelez la méthode `event.completed({allowEvent:Boolean})`.</span><span class="sxs-lookup"><span data-stu-id="7d941-432">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="7d941-433">La propriété `allowEvent` est une valeur booléenne.</span><span class="sxs-lookup"><span data-stu-id="7d941-433">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="7d941-434">Si la valeur est définie sur `true`, l’envoi est autorisé.</span><span class="sxs-lookup"><span data-stu-id="7d941-434">If set to `true`, send is allowed.</span></span> <span data-ttu-id="7d941-435">Si la valeur est définie sur `false`, l’envoi du message est bloqué.</span><span class="sxs-lookup"><span data-stu-id="7d941-435">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="7d941-436">Pour plus d’informations, consultez les articles relatifs à [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et à [completed](/javascript/api/office/office.addincommands.event).</span><span class="sxs-lookup"><span data-stu-id="7d941-436">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="7d941-437">Méthodes `replaceAsync`, `removeAsync` et `getAllAsync`</span><span class="sxs-lookup"><span data-stu-id="7d941-437">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="7d941-438">En plus de la méthode `addAsync`, l'objet `NotificationMessages` inclut également les méthodes `replaceAsync`, `removeAsync` et `getAllAsync`.</span><span class="sxs-lookup"><span data-stu-id="7d941-438">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="7d941-439">Ces méthodes ne sont pas utilisées dans cet exemple de code.</span><span class="sxs-lookup"><span data-stu-id="7d941-439">These methods are not used in this code sample.</span></span>  <span data-ttu-id="7d941-440">Pour plus d’informations, consultez l’article relatif à l’objet [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span><span class="sxs-lookup"><span data-stu-id="7d941-440">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="7d941-441">Code vérificateur de l’objet et de la ligne CC</span><span class="sxs-lookup"><span data-stu-id="7d941-441">Subject and CC checker code</span></span>

<span data-ttu-id="7d941-442">L’exemple de code suivant vous montre comment ajouter un destinataire à la ligne Cc et vérifier que le message comporte un objet pendant l’envoi.</span><span class="sxs-lookup"><span data-stu-id="7d941-442">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="7d941-443">Cet exemple utilise la fonctionnalité d’envoi pour autoriser ou interdire l’envoi d’un e-mail.</span><span class="sxs-lookup"><span data-stu-id="7d941-443">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

```js
// Invoke by Contoso Subject and CC Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateSubjectAndCC(event) {
    shouldChangeSubjectOnSend(event);
}

// Determine whether the subject should be changed. If it is already changed, allow send. Otherwise change it.
// <param name="event">ItemSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
    mailboxItem.subject.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            addCCOnSend(asyncResult.asyncContext);
            //console.log(asyncResult.value);
            // Match string.
            var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
            // Add [Checked]: to subject line.
            subject = '[Checked]: ' + asyncResult.value;

            // Determine whether a string is blank, null, or undefined.
            // If yes, block send and display information bar to notify sender to add a subject.
            if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                if (!checkSubject) {
                    subjectOnSendChange(subject, asyncResult.asyncContext);
                    //console.log(checkSubject);
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }
            }
        });
}

// Add a CC to the email. In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">ItemSend event passed from calling function</param>
function addCCOnSend(event) {
    mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });
}

// Determine whether the subject should be changed. If it is already changed, allow send, otherwise change it.
// <param name="subject">Subject to set.</param>
// <param name="event">ItemSend event passed from the calling function.</param>
function subjectOnSendChange(subject, event) {
    mailboxItem.subject.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                // Block send.
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // Allow send.
                asyncResult.asyncContext.completed({ allowEvent: true });
            }
        });
}
```

<span data-ttu-id="7d941-p156">Pour savoir comment ajouter un destinataire à la ligne Cc et vérifier que le message comporte une ligne d’objet pendant l’envoi, et découvrir les API disponibles, consultez l’article relatif à l’exemple [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send). Le code est accompagné de commentaires détaillés.</span><span class="sxs-lookup"><span data-stu-id="7d941-p156">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="7d941-446">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7d941-446">See also</span></span>

- [<span data-ttu-id="7d941-447">Présentation de l’architecture et des fonctionnalités des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="7d941-447">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="7d941-448">Démonstration de la commande du complément Outlook</span><span class="sxs-lookup"><span data-stu-id="7d941-448">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)