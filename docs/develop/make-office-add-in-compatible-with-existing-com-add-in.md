---
title: Rendre votre complément Office compatible avec un complément COM existant
description: Activer la compatibilité entre votre complément Office et un complément COM équivalent
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 3577b8fe4b4a26ac5d0af85cc5c2f96a7a8dc010
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128051"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a><span data-ttu-id="dbc08-103">Faire en sorte que votre complément Office soit compatible avec un complément COM existant (aperçu)</span><span class="sxs-lookup"><span data-stu-id="dbc08-103">Make your Office Add-in compatible with an existing COM add-in (preview)</span></span>

<span data-ttu-id="dbc08-104">Si vous disposez d’un complément COM existant, vous pouvez créer une fonctionnalité équivalente dans votre complément Office, ce qui permet à votre solution de s’exécuter sur d’autres plateformes, telles qu’Office sur le Web ou Office sur Mac.</span><span class="sxs-lookup"><span data-stu-id="dbc08-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Office on Mac.</span></span> <span data-ttu-id="dbc08-105">Dans certains cas, votre complément Office peut ne pas être en mesure de fournir toutes les fonctionnalités disponibles dans le complément COM correspondant.</span><span class="sxs-lookup"><span data-stu-id="dbc08-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="dbc08-106">Dans ce cas, votre complément COM peut fournir une meilleure expérience utilisateur sur Windows que le complément Office correspondant.</span><span class="sxs-lookup"><span data-stu-id="dbc08-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="dbc08-107">Vous pouvez configurer votre complément Office de sorte que, lorsque le complément COM équivalent est déjà installé sur l’ordinateur d’un utilisateur, Office sur Windows exécute le complément COM au lieu du complément Office.</span><span class="sxs-lookup"><span data-stu-id="dbc08-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="dbc08-108">Le complément COM est appelé «équivalent», car Office effectuera une transition transparente entre le complément COM et le complément Office en fonction de celui sur lequel est installé l’ordinateur d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dbc08-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="dbc08-109">Cette fonctionnalité est actuellement en préversion et n’est pas prise en charge dans les environnements de production.</span><span class="sxs-lookup"><span data-stu-id="dbc08-109">This feature is currently in preview and not supported for use in production environments.</span></span> <span data-ttu-id="dbc08-110">Elle est disponible dans Excel, Word et PowerPoint version 16.0.11629.20214 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="dbc08-110">It's available in Excel, Word, and PowerPoint version 16.0.11629.20214 or later.</span></span> <span data-ttu-id="dbc08-111">Pour accéder à cette version, vous devez disposer d’un abonnement Office 365 et rejoindre le programme [Office](https://products.office.com/office-insider) Insider au niveau Insider. \*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="dbc08-111">To access this build, you must have an Office 365 subscription and join the [Office Insider](https://products.office.com/office-insider) program at the **Insider** level.</span></span>

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="dbc08-112">Spécifier un complément COM équivalent dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="dbc08-112">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="dbc08-113">Pour activer la compatibilité entre votre complément Office et le complément COM, identifiez le complément COM équivalent dans le [manifeste](add-in-manifests.md) de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="dbc08-113">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="dbc08-114">Office sur Windows utilisera ensuite le complément COM au lieu du complément Office, s’ils sont tous les deux installés.</span><span class="sxs-lookup"><span data-stu-id="dbc08-114">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="dbc08-115">L’exemple suivant montre la partie du manifeste qui spécifie un complément COM sous la forme d’un complément équivalent.</span><span class="sxs-lookup"><span data-stu-id="dbc08-115">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="dbc08-116">La valeur de l' `ProgId` élément identifie le complément COM et l' `EquivalentAddins` élément doit être placé immédiatement avant la balise de `VersionOverrides` fermeture.</span><span class="sxs-lookup"><span data-stu-id="dbc08-116">The value of the `ProgId` element identifies the COM add-in and the `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  <EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="dbc08-117">Pour plus d’informations sur les compléments COM et la compatibilité des FDU XLL, consultez [la rubrique faire en sorte que les fonctions personnalisées soient compatibles avec les fonctions XLL définies par l’utilisateur](../excel/make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="dbc08-117">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="dbc08-118">Comportement équivalent pour les utilisateurs</span><span class="sxs-lookup"><span data-stu-id="dbc08-118">Equivalent behavior for users</span></span>

<span data-ttu-id="dbc08-119">Lorsqu’un complément COM équivalent est spécifié dans le manifeste du complément Office, Office sur Windows n’affiche pas l’interface utilisateur (IU) de votre complément Office si le complément COM équivalent est installé.</span><span class="sxs-lookup"><span data-stu-id="dbc08-119">When an equivalent COM add-in is specified in the Office Add-in manifest, Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="dbc08-120">Office masque uniquement les boutons du ruban du complément Office et n’empêche pas l’installation.</span><span class="sxs-lookup"><span data-stu-id="dbc08-120">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="dbc08-121">Par conséquent, votre complément Office continuera à apparaître aux emplacements suivants au sein de l’interface utilisateur:</span><span class="sxs-lookup"><span data-stu-id="dbc08-121">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="dbc08-122">Sous **mes compléments**</span><span class="sxs-lookup"><span data-stu-id="dbc08-122">Under **My add-ins**</span></span>
- <span data-ttu-id="dbc08-123">Comme entrée dans le gestionnaire de ruban</span><span class="sxs-lookup"><span data-stu-id="dbc08-123">As an entry in the ribbon manager</span></span>

> [!NOTE]
> <span data-ttu-id="dbc08-124">La spécification d’un complément COM équivalent dans le manifeste n’a aucun effet sur les autres plateformes, comme Office sur le Web ou Mac.</span><span class="sxs-lookup"><span data-stu-id="dbc08-124">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or Mac.</span></span>

<span data-ttu-id="dbc08-125">Les scénarios suivants décrivent ce qui se produit en fonction de la manière dont l’utilisateur acquiert le complément Office.</span><span class="sxs-lookup"><span data-stu-id="dbc08-125">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="dbc08-126">AppSource acquisition d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="dbc08-126">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="dbc08-127">Si un utilisateur acquiert le complément Office à partir de AppSource et que le complément COM équivalent est déjà installé, Office:</span><span class="sxs-lookup"><span data-stu-id="dbc08-127">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="dbc08-128">Installez le complément Office.</span><span class="sxs-lookup"><span data-stu-id="dbc08-128">Install the Office Add-in.</span></span>
2. <span data-ttu-id="dbc08-129">Masquer l’interface utilisateur du complément Office dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="dbc08-129">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="dbc08-130">Afficher un appel pour l’utilisateur qui pointe vers le bouton du ruban de complément COM.</span><span class="sxs-lookup"><span data-stu-id="dbc08-130">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="dbc08-131">Déploiement centralisé du complément Office</span><span class="sxs-lookup"><span data-stu-id="dbc08-131">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="dbc08-132">Si un administrateur déploie le complément Office sur son client à l’aide d’un déploiement centralisé, et que le complément COM équivalent est déjà installé, l’utilisateur doit redémarrer Office avant de voir les modifications.</span><span class="sxs-lookup"><span data-stu-id="dbc08-132">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="dbc08-133">Après le redémarrage d’Office, il peut:</span><span class="sxs-lookup"><span data-stu-id="dbc08-133">After Office restarts, it will:</span></span>

1. <span data-ttu-id="dbc08-134">Installez le complément Office.</span><span class="sxs-lookup"><span data-stu-id="dbc08-134">Install the Office Add-in.</span></span>
2. <span data-ttu-id="dbc08-135">Masquer l’interface utilisateur du complément Office dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="dbc08-135">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="dbc08-136">Afficher un appel pour l’utilisateur qui pointe vers le bouton du ruban de complément COM.</span><span class="sxs-lookup"><span data-stu-id="dbc08-136">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="dbc08-137">Document partagé avec un complément Office incorporé</span><span class="sxs-lookup"><span data-stu-id="dbc08-137">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="dbc08-138">Si un utilisateur a installé le complément COM, puis qu’il obtient un document partagé avec le complément Office incorporé, lorsqu’il ouvre le document, Office:</span><span class="sxs-lookup"><span data-stu-id="dbc08-138">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="dbc08-139">Inviter l’utilisateur à approuver le complément Office.</span><span class="sxs-lookup"><span data-stu-id="dbc08-139">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="dbc08-140">S’il est approuvé, le complément Office est installé.</span><span class="sxs-lookup"><span data-stu-id="dbc08-140">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="dbc08-141">Masquer l’interface utilisateur du complément Office dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="dbc08-141">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="dbc08-142">Autre comportement de complément COM</span><span class="sxs-lookup"><span data-stu-id="dbc08-142">Other COM add-in behavior</span></span>

<span data-ttu-id="dbc08-143">Si un utilisateur désinstalle le complément COM équivalent, Office sur Windows restaure l’interface utilisateur du complément Office.</span><span class="sxs-lookup"><span data-stu-id="dbc08-143">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="dbc08-144">Une fois que vous avez spécifié un complément COM équivalent pour votre complément Office, Office cesse de traiter les mises à jour pour votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="dbc08-144">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="dbc08-145">Pour obtenir les dernières mises à jour pour le complément Office, l’utilisateur doit d’abord désinstaller le complément COM.</span><span class="sxs-lookup"><span data-stu-id="dbc08-145">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="dbc08-146">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="dbc08-146">See also</span></span>

- [<span data-ttu-id="dbc08-147">Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="dbc08-147">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
