---
title: Rendre votre complément Office compatible avec un complément COM existant
description: Activez la compatibilité entre votre compl?ment Office et un compl?ment COM équivalent.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: b5235255987cc6a644491bc548b92701b350a179
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836851"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="33fbd-103">Rendre votre complément Office compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="33fbd-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="33fbd-104">Si vous avez un compl?ment COM existant, vous pouvez créer des fonctionnalités équivalentes dans votre compl?ment Office, ce qui permet à votre solution de s’exécuter sur d’autres plateformes telles qu’Office sur le web ou Mac.</span><span class="sxs-lookup"><span data-stu-id="33fbd-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac.</span></span> <span data-ttu-id="33fbd-105">Dans certains cas, il se peut que votre compl?ment Office ne soit pas en mesure de fournir toutes les fonctionnalités disponibles dans le compl?ment COM correspondant.</span><span class="sxs-lookup"><span data-stu-id="33fbd-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="33fbd-106">Dans ces situations, votre compl?ment COM peut fournir une meilleure expérience utilisateur sur Windows que le compl?ment Office correspondant.</span><span class="sxs-lookup"><span data-stu-id="33fbd-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="33fbd-107">Vous pouvez configurer votre compl?ment Office de sorte que lorsque le compl?ment COM équivalent est déjà install sur l’ordinateur d’un utilisateur, Office sur Windows exécute le compl?ment COM au lieu du compl?ment Office.</span><span class="sxs-lookup"><span data-stu-id="33fbd-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="33fbd-108">Le add-in COM est appelé « équivalent », car Office passe en toute transparence entre le compl?ment COM et le compl?ment Office en fonction de l’ordinateur d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="33fbd-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="33fbd-109">Cette fonctionnalité est prise en charge par les plateformes suivantes, lorsqu’elle est connectée à un abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="33fbd-109">This feature is supported by the following platforms, when connected to a Microsoft 365 subscription.</span></span>
>
> - <span data-ttu-id="33fbd-110">Excel, Word et PowerPoint sur le web</span><span class="sxs-lookup"><span data-stu-id="33fbd-110">Excel, Word, and PowerPoint on the web</span></span>
> - <span data-ttu-id="33fbd-111">Excel, Word et PowerPoint sur Windows (version 1904 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="33fbd-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="33fbd-112">Excel, Word et PowerPoint sur Mac (version 13.329 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="33fbd-112">Excel, Word, and PowerPoint on Mac (version 13.329 or later)</span></span>
> - <span data-ttu-id="33fbd-113">Outlook sur Windows (version 2102 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="33fbd-113">Outlook on Windows (version 2102 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in"></a><span data-ttu-id="33fbd-114">Spécifier un compl?ment COM équivalent</span><span class="sxs-lookup"><span data-stu-id="33fbd-114">Specify an equivalent COM add-in</span></span>

### <a name="manifest"></a><span data-ttu-id="33fbd-115">Manifeste</span><span class="sxs-lookup"><span data-stu-id="33fbd-115">Manifest</span></span>

> [!IMPORTANT]
> <span data-ttu-id="33fbd-116">S’applique à Excel, PowerPoint et Word.</span><span class="sxs-lookup"><span data-stu-id="33fbd-116">Applies to Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="33fbd-117">Prise en charge d’Outlook bientôt disponible.</span><span class="sxs-lookup"><span data-stu-id="33fbd-117">Outlook support coming soon.</span></span>

<span data-ttu-id="33fbd-118">Pour activer la compatibilité entre votre compl?ment Office et votre compl?ment COM, identifiez le compl?ment COM équivalent dans le manifeste de votre compl?ment Office. [](add-in-manifests.md)</span><span class="sxs-lookup"><span data-stu-id="33fbd-118">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="33fbd-119">Ensuite, Office sur Windows utilisera le compl?ment COM au lieu du compl?ment Office, s’ils sont tous deux install s.</span><span class="sxs-lookup"><span data-stu-id="33fbd-119">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="33fbd-120">L’exemple suivant montre la partie du manifeste qui spécifie un compl?ment COM en tant que compl?ment équivalent.</span><span class="sxs-lookup"><span data-stu-id="33fbd-120">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="33fbd-121">La valeur de l’élément identifie le add-in COM et l’élément `ProgId` [EquivalentAddins](../reference/manifest/equivalentaddins.md) doit être placé immédiatement avant la balise `VersionOverrides` de fermeture.</span><span class="sxs-lookup"><span data-stu-id="33fbd-121">The value of the `ProgId` element identifies the COM add-in and the [EquivalentAddins](../reference/manifest/equivalentaddins.md) element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="33fbd-122">Pour plus d’informations sur le module complémentaire COM et la compatibilité XLL UDF, voir Rendre vos fonctions personnalisées compatibles avec les fonctions [XLL définies par l’utilisateur.](../excel/make-custom-functions-compatible-with-xll-udf.md)</span><span class="sxs-lookup"><span data-stu-id="33fbd-122">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

### <a name="group-policy"></a><span data-ttu-id="33fbd-123">Stratégie de groupe</span><span class="sxs-lookup"><span data-stu-id="33fbd-123">Group policy</span></span>

> [!IMPORTANT]
> <span data-ttu-id="33fbd-124">S’applique uniquement à Outlook.</span><span class="sxs-lookup"><span data-stu-id="33fbd-124">Applies to Outlook only.</span></span>

<span data-ttu-id="33fbd-125">Pour déclarer la compatibilité entre votre compl?ment web Outlook et le compl?ment COM/VSTO, identifiez le compl?ment COM équivalent dans la stratégie de groupe Deactiver les compl?ments web Outlook dont les compl?ments COM ou **VSTO équivalents** sont install s en configurant sur l’ordinateur de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="33fbd-125">To declare compatibility between your Outlook web add-in and COM/VSTO add-in, identify the equivalent COM add-in in the group policy **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** by configuring on the user's machine.</span></span> <span data-ttu-id="33fbd-126">Outlook sur Windows utilisera ensuite le compl?ment COM au lieu du compl?ment web, s’ils sont tous deux install s.</span><span class="sxs-lookup"><span data-stu-id="33fbd-126">Then Outlook on Windows will use the COM add-in instead of the web add-in, if they're both installed.</span></span>

1. <span data-ttu-id="33fbd-127">Téléchargez le dernier [outil Modèles d’administration,](https://www.microsoft.com/download/details.aspx?id=49030)en vous important des instructions d’installation **de l’outil.**</span><span class="sxs-lookup"><span data-stu-id="33fbd-127">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.</span></span>
1. <span data-ttu-id="33fbd-128">Ouvrez l’Éditeur de stratégie de groupe local (**gpedit.msc**).</span><span class="sxs-lookup"><span data-stu-id="33fbd-128">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="33fbd-129">Accédez **à Modèles** d’administration de configuration  >     >  **utilisateur Microsoft Outlook 2016**  >  **Divers.**</span><span class="sxs-lookup"><span data-stu-id="33fbd-129">Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.</span></span>
1. <span data-ttu-id="33fbd-130">Sélectionnez le paramètre Désactiver les **compl?ments web Outlook** dont les compl?ments COM ou VSTO équivalents sont install s .</span><span class="sxs-lookup"><span data-stu-id="33fbd-130">Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.</span></span>
1. <span data-ttu-id="33fbd-131">Ouvrez le lien pour modifier le paramètre de stratégie.</span><span class="sxs-lookup"><span data-stu-id="33fbd-131">Open the link to edit the policy setting.</span></span>
1. <span data-ttu-id="33fbd-132">Dans la boîte **de dialogue, les applications web Outlook** sont à désactiver :</span><span class="sxs-lookup"><span data-stu-id="33fbd-132">In the dialog **Outlook web add-ins to deactivate**:</span></span>
    1. <span data-ttu-id="33fbd-133">Définissez **le nom de** la valeur sur la valeur trouvée dans le manifeste du `Id` add-in web.</span><span class="sxs-lookup"><span data-stu-id="33fbd-133">Set **Value name** to the `Id` found in the web add-in's manifest.</span></span> <span data-ttu-id="33fbd-134">**Important**: *n’ajoutez* pas d’accolades `{}` autour de l’entrée.</span><span class="sxs-lookup"><span data-stu-id="33fbd-134">**Important**: Do *not* add curly braces `{}` around the entry.</span></span>
    1. <span data-ttu-id="33fbd-135">Définissez **la** valeur sur la valeur du `ProgId` compl?ment COM/VSTO équivalent.</span><span class="sxs-lookup"><span data-stu-id="33fbd-135">Set **Value** to the `ProgId` of the equivalent COM/VSTO add-in.</span></span>
    1. <span data-ttu-id="33fbd-136">Sélectionnez **OK** pour mettre la mise à jour en vigueur.</span><span class="sxs-lookup"><span data-stu-id="33fbd-136">Select **OK** to put the update into effect.</span></span>
    <span data-ttu-id="33fbd-137">![Capture d’écran montrant la boîte de dialogue « Les applications web Outlook à désactiver »](../images/outlook-deactivate-gpo-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="33fbd-137">![Screenshot showing the dialog "Outlook web add-ins to deactivate"](../images/outlook-deactivate-gpo-dialog.png)</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="33fbd-138">Comportement équivalent pour les utilisateurs</span><span class="sxs-lookup"><span data-stu-id="33fbd-138">Equivalent behavior for users</span></span>

<span data-ttu-id="33fbd-139">Lorsqu’un compl?ment [COM](#specify-an-equivalent-com-add-in)équivalent est spécifié, Office sur Windows n’affiche pas l’interface utilisateur de votre compl?ment Office si le compl?ment COM équivalent est install .</span><span class="sxs-lookup"><span data-stu-id="33fbd-139">When an [equivalent COM add-in is specified](#specify-an-equivalent-com-add-in), Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="33fbd-140">Office masque uniquement les boutons du ruban du add-in Office et n’empêche pas l’installation.</span><span class="sxs-lookup"><span data-stu-id="33fbd-140">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="33fbd-141">Par conséquent, votre add-in Office apparaîtra toujours aux emplacements suivants dans l’interface utilisateur :</span><span class="sxs-lookup"><span data-stu-id="33fbd-141">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="33fbd-142">Sous **Mes modules**</span><span class="sxs-lookup"><span data-stu-id="33fbd-142">Under **My add-ins**</span></span>
- <span data-ttu-id="33fbd-143">En tant qu’entrée dans le gestionnaire du ruban (Excel, Word et PowerPoint uniquement)</span><span class="sxs-lookup"><span data-stu-id="33fbd-143">As an entry in the ribbon manager (Excel, Word, and PowerPoint only)</span></span>

> [!NOTE]
> <span data-ttu-id="33fbd-144">La spécification d’un équivalent com dans le manifeste n’a aucun effet sur les autres plateformes telles qu’Office sur le web ou sur Mac.</span><span class="sxs-lookup"><span data-stu-id="33fbd-144">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or on Mac.</span></span>

<span data-ttu-id="33fbd-145">Les scénarios suivants décrivent ce qui se produit en fonction de la façon dont l’utilisateur acquiert le add-in Office.</span><span class="sxs-lookup"><span data-stu-id="33fbd-145">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="33fbd-146">Acquisition d’un add-in Office dans AppSource</span><span class="sxs-lookup"><span data-stu-id="33fbd-146">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="33fbd-147">Si un utilisateur acquiert le compl?ment Office auprès d’AppSource et que le compl?ment COM équivalent est déjà install ? , Office :</span><span class="sxs-lookup"><span data-stu-id="33fbd-147">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="33fbd-148">Installez le add-in Office.</span><span class="sxs-lookup"><span data-stu-id="33fbd-148">Install the Office Add-in.</span></span>
2. <span data-ttu-id="33fbd-149">Masquer l’interface utilisateur du add-in Office dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="33fbd-149">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="33fbd-150">Affichez un appel pour l’utilisateur qui pointe sur le bouton du ruban du compl?ment COM.</span><span class="sxs-lookup"><span data-stu-id="33fbd-150">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="33fbd-151">Déploiement centralisé d’un add-in Office</span><span class="sxs-lookup"><span data-stu-id="33fbd-151">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="33fbd-152">Si un administrateur déploie le add-in Office sur son client à l’aide d’un déploiement centralisé et que le module com équivalent est déjà installé, l’utilisateur doit redémarrer Office avant de voir des modifications.</span><span class="sxs-lookup"><span data-stu-id="33fbd-152">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="33fbd-153">Après le redémarrage d’Office, il :</span><span class="sxs-lookup"><span data-stu-id="33fbd-153">After Office restarts, it will:</span></span>

1. <span data-ttu-id="33fbd-154">Installez le add-in Office.</span><span class="sxs-lookup"><span data-stu-id="33fbd-154">Install the Office Add-in.</span></span>
2. <span data-ttu-id="33fbd-155">Masquer l’interface utilisateur du add-in Office dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="33fbd-155">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="33fbd-156">Affichez un appel pour l’utilisateur qui pointe sur le bouton du ruban du compl?ment COM.</span><span class="sxs-lookup"><span data-stu-id="33fbd-156">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="33fbd-157">Document partagé avec un add-in Office incorporé</span><span class="sxs-lookup"><span data-stu-id="33fbd-157">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="33fbd-158">Si un utilisateur a installé le compl?ment COM, puis obtient un document partagé avec le compl?ment Office incorporé, alors lorsqu’il ouvre le document, Office :</span><span class="sxs-lookup"><span data-stu-id="33fbd-158">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="33fbd-159">Invitez l’utilisateur à faire confiance au add-in Office.</span><span class="sxs-lookup"><span data-stu-id="33fbd-159">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="33fbd-160">S’il est approuvé, le add-in Office s’installe.</span><span class="sxs-lookup"><span data-stu-id="33fbd-160">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="33fbd-161">Masquer l’interface utilisateur du add-in Office dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="33fbd-161">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="33fbd-162">Comportement des autres compl?ments COM</span><span class="sxs-lookup"><span data-stu-id="33fbd-162">Other COM add-in behavior</span></span>

### <a name="excel-powerpoint-word"></a><span data-ttu-id="33fbd-163">Excel, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="33fbd-163">Excel, PowerPoint, Word</span></span>

<span data-ttu-id="33fbd-164">Si un utilisateur désinstalle le compl?ment COM équivalent, Office sur Windows restaure l’interface utilisateur du compl?ment Office.</span><span class="sxs-lookup"><span data-stu-id="33fbd-164">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="33fbd-165">Après avoir spécifié un compl?ment COM équivalent pour votre compl?ment Office, Office cesse de traiter les mises à jour pour votre compl?ment Office.</span><span class="sxs-lookup"><span data-stu-id="33fbd-165">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="33fbd-166">Pour obtenir les dernières mises à jour pour le compl?ment Office, l’utilisateur doit d’abord désinstaller le compl?ment COM.</span><span class="sxs-lookup"><span data-stu-id="33fbd-166">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

### <a name="outlook"></a><span data-ttu-id="33fbd-167">Outlook</span><span class="sxs-lookup"><span data-stu-id="33fbd-167">Outlook</span></span>

<span data-ttu-id="33fbd-168">Le add-in COM/VSTO doit être connecté au moment du début d’Outlook afin que le compl?ment web correspondant soit désactivé.</span><span class="sxs-lookup"><span data-stu-id="33fbd-168">The COM/VSTO add-in must be connected when Outlook is started in order for the corresponding web add-in to be disabled.</span></span>

<span data-ttu-id="33fbd-169">Si le compl?ment COM/VSTO est alors déconnecté lors d’une session Outlook suivante, le compl?ment web restera probablement désactivé jusqu’au redémarrage d’Outlook.</span><span class="sxs-lookup"><span data-stu-id="33fbd-169">If the COM/VSTO add-in is then disconnected during a subsequent Outlook session, the web add-in will likely remain disabled until Outlook is restarted.</span></span>

## <a name="see-also"></a><span data-ttu-id="33fbd-170">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="33fbd-170">See also</span></span>

- [<span data-ttu-id="33fbd-171">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="33fbd-171">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
