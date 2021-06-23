---
title: Rendre votre complément Office compatible avec un complément COM existant
description: Activez la compatibilité entre votre Office et votre équivalent COM.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: e2ab1bb1eda548ff8e0923b8fbccfa9e007a6a0c
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075998"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="c72f2-103">Rendre votre complément Office compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="c72f2-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="c72f2-104">Si vous avez un compl?ment COM existant, vous pouvez créer des fonctionnalités équivalentes dans votre compl?ment Office, permettant ainsi votre solution de s’exécuter sur d’autres plateformes telles que Office sur le Web ou Mac.</span><span class="sxs-lookup"><span data-stu-id="c72f2-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac.</span></span> <span data-ttu-id="c72f2-105">Dans certains cas, votre Office peut ne pas être en mesure de fournir toutes les fonctionnalités disponibles dans le compl?ment COM correspondant.</span><span class="sxs-lookup"><span data-stu-id="c72f2-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="c72f2-106">Dans ces situations, votre compl?ment COM peut fournir une meilleure expérience utilisateur sur Windows que l’interface Office compl?ments peut fournir.</span><span class="sxs-lookup"><span data-stu-id="c72f2-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="c72f2-107">Vous pouvez configurer votre compl?ment Office de sorte que lorsque le compl?ment COM équivalent est déjà install sur l’ordinateur d’un utilisateur, Office sur Windows exécute le compl?ment COM au lieu du compl?ment Office.</span><span class="sxs-lookup"><span data-stu-id="c72f2-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="c72f2-108">Le add-in COM est appelé « équivalent », car Office passe en toute transparence entre le compl?ment COM et le compl?ment Office en fonction de celui qui est install ? l’ordinateur d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c72f2-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="c72f2-109">Cette fonctionnalité est prise en charge par les plateformes suivantes, lorsqu’elles sont connectées à Microsoft 365 abonnement.</span><span class="sxs-lookup"><span data-stu-id="c72f2-109">This feature is supported by the following platforms, when connected to a Microsoft 365 subscription.</span></span>
>
> - <span data-ttu-id="c72f2-110">Excel, Word et PowerPoint sur le web</span><span class="sxs-lookup"><span data-stu-id="c72f2-110">Excel, Word, and PowerPoint on the web</span></span>
> - <span data-ttu-id="c72f2-111">Excel, Word et PowerPoint sur Windows (version 1904 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="c72f2-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="c72f2-112">Excel, Word et PowerPoint mac (version 13.329 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="c72f2-112">Excel, Word, and PowerPoint on Mac (version 13.329 or later)</span></span>
> - <span data-ttu-id="c72f2-113">Outlook sur Windows (version 2102 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="c72f2-113">Outlook on Windows (version 2102 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in"></a><span data-ttu-id="c72f2-114">Spécifier un compl?ment COM équivalent</span><span class="sxs-lookup"><span data-stu-id="c72f2-114">Specify an equivalent COM add-in</span></span>

### <a name="manifest"></a><span data-ttu-id="c72f2-115">Manifeste</span><span class="sxs-lookup"><span data-stu-id="c72f2-115">Manifest</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c72f2-116">S’applique Excel, PowerPoint et Word.</span><span class="sxs-lookup"><span data-stu-id="c72f2-116">Applies to Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="c72f2-117">Outlook prise en charge sera bientôt disponible.</span><span class="sxs-lookup"><span data-stu-id="c72f2-117">Outlook support coming soon.</span></span>

<span data-ttu-id="c72f2-118">Pour activer la compatibilité entre votre Office et votre compl?ment COM, identifiez [](add-in-manifests.md) le compl?ment COM équivalent dans le manifeste de votre Office compl?ment.</span><span class="sxs-lookup"><span data-stu-id="c72f2-118">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="c72f2-119">Ensuite, Office sur Windows utilisera le compl?ment COM au lieu du compl?ment Office, s’ils sont tous les deux install s.</span><span class="sxs-lookup"><span data-stu-id="c72f2-119">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="c72f2-120">L’exemple suivant montre la partie du manifeste qui spécifie un compl?ment COM en tant que compl?ment équivalent.</span><span class="sxs-lookup"><span data-stu-id="c72f2-120">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="c72f2-121">La valeur de l’élément identifie le add-in COM et l’élément `ProgId` [EquivalentAddins](../reference/manifest/equivalentaddins.md) doit être placé immédiatement avant la balise `VersionOverrides` de fermeture.</span><span class="sxs-lookup"><span data-stu-id="c72f2-121">The value of the `ProgId` element identifies the COM add-in and the [EquivalentAddins](../reference/manifest/equivalentaddins.md) element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

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
> <span data-ttu-id="c72f2-122">Pour plus d’informations sur le module complémentaire COM et la compatibilité XLL UDF, voir Rendre vos fonctions personnalisées compatibles avec les fonctions [XLL définies par l’utilisateur.](../excel/make-custom-functions-compatible-with-xll-udf.md)</span><span class="sxs-lookup"><span data-stu-id="c72f2-122">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

### <a name="group-policy"></a><span data-ttu-id="c72f2-123">Stratégie de groupe</span><span class="sxs-lookup"><span data-stu-id="c72f2-123">Group policy</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c72f2-124">S’applique Outlook uniquement.</span><span class="sxs-lookup"><span data-stu-id="c72f2-124">Applies to Outlook only.</span></span>

<span data-ttu-id="c72f2-125">Pour déclarer la compatibilité entre votre compl?ment web Outlook et le compl?ment COM/VSTO, identifiez le compl?ment COM équivalent dans la stratégie de groupe **Deactiver** les compl?ments web Outlook dont le compl?ment COM ou VSTO équivalent est install s en configurant sur l’ordinateur de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c72f2-125">To declare compatibility between your Outlook web add-in and COM/VSTO add-in, identify the equivalent COM add-in in the group policy **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** by configuring on the user's machine.</span></span> <span data-ttu-id="c72f2-126">Ensuite, Outlook sur Windows utilisera le compl?ment COM au lieu du compl?ment web, s’ils sont tous deux install s.</span><span class="sxs-lookup"><span data-stu-id="c72f2-126">Then Outlook on Windows will use the COM add-in instead of the web add-in, if they're both installed.</span></span>

1. <span data-ttu-id="c72f2-127">Téléchargez le dernier [outil Modèles d’administration,](https://www.microsoft.com/download/details.aspx?id=49030)en vous important des instructions d’installation **de l’outil.**</span><span class="sxs-lookup"><span data-stu-id="c72f2-127">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.</span></span>
1. <span data-ttu-id="c72f2-128">Ouvrez l’Éditeur de stratégie de groupe local (**gpedit.msc**).</span><span class="sxs-lookup"><span data-stu-id="c72f2-128">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="c72f2-129">Accédez **aux**  >  **modèles d’administration** de configuration utilisateur   >  **Microsoft Outlook 2016**  >  **divers.**</span><span class="sxs-lookup"><span data-stu-id="c72f2-129">Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.</span></span>
1. <span data-ttu-id="c72f2-130">Sélectionnez le paramètre Désactiver Outlook de sites web dont l’équivalent **COM ou VSTO est installé.**</span><span class="sxs-lookup"><span data-stu-id="c72f2-130">Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.</span></span>
1. <span data-ttu-id="c72f2-131">Ouvrez le lien pour modifier le paramètre de stratégie.</span><span class="sxs-lookup"><span data-stu-id="c72f2-131">Open the link to edit the policy setting.</span></span>
1. <span data-ttu-id="c72f2-132">Dans la boîte **de dialogue Outlook les** applications web à désactiver :</span><span class="sxs-lookup"><span data-stu-id="c72f2-132">In the dialog **Outlook web add-ins to deactivate**:</span></span>
    1. <span data-ttu-id="c72f2-133">Définissez **le nom de** la valeur sur la valeur trouvée dans le manifeste du `Id` add-in web.</span><span class="sxs-lookup"><span data-stu-id="c72f2-133">Set **Value name** to the `Id` found in the web add-in's manifest.</span></span> <span data-ttu-id="c72f2-134">**Important**: *n’ajoutez* pas d’accolades `{}` autour de l’entrée.</span><span class="sxs-lookup"><span data-stu-id="c72f2-134">**Important**: Do *not* add curly braces `{}` around the entry.</span></span>
    1. <span data-ttu-id="c72f2-135">Définissez **la** valeur sur la valeur du VSTO `ProgId` com/VSTO équivalent.</span><span class="sxs-lookup"><span data-stu-id="c72f2-135">Set **Value** to the `ProgId` of the equivalent COM/VSTO add-in.</span></span>
    1. <span data-ttu-id="c72f2-136">Sélectionnez **OK** pour mettre la mise à jour en vigueur.</span><span class="sxs-lookup"><span data-stu-id="c72f2-136">Select **OK** to put the update into effect.</span></span>
    <span data-ttu-id="c72f2-137">![Screenshot showing the dialog « Outlook web add-ins to deactivate ».](../images/outlook-deactivate-gpo-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="c72f2-137">![Screenshot showing the dialog "Outlook web add-ins to deactivate".](../images/outlook-deactivate-gpo-dialog.png)</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="c72f2-138">Comportement équivalent pour les utilisateurs</span><span class="sxs-lookup"><span data-stu-id="c72f2-138">Equivalent behavior for users</span></span>

<span data-ttu-id="c72f2-139">Lorsqu’un compl?ment [COM](#specify-an-equivalent-com-add-in)équivalent est spécifié, Office sur Windows n’affiche pas l’interface utilisateur de votre compl?ment Office si le compl?ment COM ex quis est install .</span><span class="sxs-lookup"><span data-stu-id="c72f2-139">When an [equivalent COM add-in is specified](#specify-an-equivalent-com-add-in), Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="c72f2-140">Office masque uniquement les boutons du ruban du Office et n’empêche pas l’installation.</span><span class="sxs-lookup"><span data-stu-id="c72f2-140">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="c72f2-141">Par conséquent, votre Office’interface utilisateur s’affiche toujours aux emplacements suivants dans l’interface utilisateur :</span><span class="sxs-lookup"><span data-stu-id="c72f2-141">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="c72f2-142">Sous **Mes modules**</span><span class="sxs-lookup"><span data-stu-id="c72f2-142">Under **My add-ins**</span></span>
- <span data-ttu-id="c72f2-143">En tant qu’entrée dans le gestionnaire du ruban (Excel, Word et PowerPoint uniquement)</span><span class="sxs-lookup"><span data-stu-id="c72f2-143">As an entry in the ribbon manager (Excel, Word, and PowerPoint only)</span></span>

> [!NOTE]
> <span data-ttu-id="c72f2-144">La spécification d’un module com équivalent dans le manifeste n’a aucun effet sur les autres plateformes telles que Office sur le Web ou Mac.</span><span class="sxs-lookup"><span data-stu-id="c72f2-144">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or on Mac.</span></span>

<span data-ttu-id="c72f2-145">Les scénarios suivants décrivent ce qui se produit en fonction de la façon dont l’utilisateur acquiert le Office de contenu.</span><span class="sxs-lookup"><span data-stu-id="c72f2-145">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="c72f2-146">Acquisition d’un Office AppSource</span><span class="sxs-lookup"><span data-stu-id="c72f2-146">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="c72f2-147">Si un utilisateur acquiert le Office à partir d’AppSource et que le module com équivalent est déjà installé, Office :</span><span class="sxs-lookup"><span data-stu-id="c72f2-147">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="c72f2-148">Installez le Office’installation.</span><span class="sxs-lookup"><span data-stu-id="c72f2-148">Install the Office Add-in.</span></span>
2. <span data-ttu-id="c72f2-149">Masquez l’Office’interface utilisateur du add-in dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="c72f2-149">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="c72f2-150">Affichez un appel pour l’utilisateur qui pointe sur le bouton du ruban du compl?ment COM.</span><span class="sxs-lookup"><span data-stu-id="c72f2-150">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="c72f2-151">Déploiement centralisé du Office de bureau</span><span class="sxs-lookup"><span data-stu-id="c72f2-151">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="c72f2-152">Si un administrateur déploie le add-in Office sur son client à l’aide d’un déploiement centralisé et que le module com équivalent est déjà installé, l’utilisateur doit redémarrer Office avant de voir les modifications.</span><span class="sxs-lookup"><span data-stu-id="c72f2-152">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="c72f2-153">Une fois Office redémarrage, il :</span><span class="sxs-lookup"><span data-stu-id="c72f2-153">After Office restarts, it will:</span></span>

1. <span data-ttu-id="c72f2-154">Installez le Office’installation.</span><span class="sxs-lookup"><span data-stu-id="c72f2-154">Install the Office Add-in.</span></span>
2. <span data-ttu-id="c72f2-155">Masquez l’Office’interface utilisateur du add-in dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="c72f2-155">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="c72f2-156">Affichez un appel pour l’utilisateur qui pointe sur le bouton du ruban du compl?ment COM.</span><span class="sxs-lookup"><span data-stu-id="c72f2-156">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="c72f2-157">Document partagé avec un Office incorporé</span><span class="sxs-lookup"><span data-stu-id="c72f2-157">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="c72f2-158">Si un utilisateur a installé le compl?ment COM, puis obtient un document partagé avec le compl?ment Office incorporé, alors lorsqu’il ouvre le document, Office :</span><span class="sxs-lookup"><span data-stu-id="c72f2-158">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="c72f2-159">Invitez l’utilisateur à faire confiance au Office de contenu.</span><span class="sxs-lookup"><span data-stu-id="c72f2-159">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="c72f2-160">S’il est approuvé, Office le module de mise en Office s’installe.</span><span class="sxs-lookup"><span data-stu-id="c72f2-160">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="c72f2-161">Masquez l’Office’interface utilisateur du add-in dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="c72f2-161">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="c72f2-162">Autre comportement des autres compl?ments COM</span><span class="sxs-lookup"><span data-stu-id="c72f2-162">Other COM add-in behavior</span></span>

### <a name="excel-powerpoint-word"></a><span data-ttu-id="c72f2-163">Excel, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="c72f2-163">Excel, PowerPoint, Word</span></span>

<span data-ttu-id="c72f2-164">Si un utilisateur désinstalle l’équivalent du compl?ment COM, Office sur Windows restaure l’interface utilisateur Office du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="c72f2-164">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="c72f2-165">Une fois que vous avez spécifié un Office COM équivalent pour votre Office, Office cesse de traiter les mises à jour de votre Office de recherche.</span><span class="sxs-lookup"><span data-stu-id="c72f2-165">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="c72f2-166">Pour obtenir les dernières mises à jour du Office, l’utilisateur doit d’abord désinstaller le compl?ment COM.</span><span class="sxs-lookup"><span data-stu-id="c72f2-166">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

### <a name="outlook"></a><span data-ttu-id="c72f2-167">Outlook</span><span class="sxs-lookup"><span data-stu-id="c72f2-167">Outlook</span></span>

<span data-ttu-id="c72f2-168">Le VSTO COM/Outlook doit être connecté pour que le module web correspondant soit désactivé.</span><span class="sxs-lookup"><span data-stu-id="c72f2-168">The COM/VSTO add-in must be connected when Outlook is started in order for the corresponding web add-in to be disabled.</span></span>

<span data-ttu-id="c72f2-169">Si le VSTO COM/VSTO est alors déconnecté au cours d’une session Outlook suivante, le compl?ment web restera probablement désactivé jusqu’au redémarrage de Outlook'</span><span class="sxs-lookup"><span data-stu-id="c72f2-169">If the COM/VSTO add-in is then disconnected during a subsequent Outlook session, the web add-in will likely remain disabled until Outlook is restarted.</span></span>

## <a name="see-also"></a><span data-ttu-id="c72f2-170">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c72f2-170">See also</span></span>

- [<span data-ttu-id="c72f2-171">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="c72f2-171">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
