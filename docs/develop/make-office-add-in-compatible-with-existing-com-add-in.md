---
title: Faire en sorte que votre complément Office soit compatible avec un complément COM existant
description: Activer la compatibilité avec un complément COM équivalent doté de la même fonctionnalité que votre complément Office
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 8f3780814163cc4dd21311b362d1d821a14b3e80
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356863"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="59084-103">Faire en sorte que votre complément Office soit compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="59084-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="59084-104">Si vous disposez d'un complément COM existant, vous pouvez créer une fonctionnalité équivalente dans votre complément Office pour étendre les fonctionnalités de votre solution à d'autres plateformes, comme Online ou macOS.</span><span class="sxs-lookup"><span data-stu-id="59084-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="59084-105">Toutefois, les compléments Office ne disposent pas de toutes les fonctionnalités disponibles dans les compléments COM. Votre complément COM peut fournir une meilleure expérience que le complément Office sur Windows dans Excel, Word et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="59084-105">However, Office Add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Office Add-in on Windows in Excel, Word, and PowerPoint.</span></span>

<span data-ttu-id="59084-106">Vous pouvez configurer votre complément Office de sorte que, lorsqu'un complément COM équivalent est déjà installé sur l'ordinateur de l'utilisateur, Office exécute le complément COM au lieu de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="59084-106">You can configure your Office Add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Office Add-in.</span></span> <span data-ttu-id="59084-107">Le complément COM est appelé «équivalent», car Office effectuera une transition transparente entre le complément COM et le complément Office en fonction de ce qui est installé sur Windows.</span><span class="sxs-lookup"><span data-stu-id="59084-107">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in depending on which is installed on Windows.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="59084-108">Spécifier un complément COM équivalent dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="59084-108">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="59084-109">Pour activer la compatibilité avec un complément COM existant, identifiez le complément COM équivalent dans le manifeste de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="59084-109">To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Office Add-in.</span></span> <span data-ttu-id="59084-110">Office utilise ensuite le complément COM au lieu de votre complément Office lors de l'exécution de Windows.</span><span class="sxs-lookup"><span data-stu-id="59084-110">Then Office will use the COM add-in instead of your Office Add-in when running on Windows.</span></span>

<span data-ttu-id="59084-111">Spécifiez `ProgID` le du complément COM équivalent.</span><span class="sxs-lookup"><span data-stu-id="59084-111">Specify the `ProgID` of the equivalent COM add-in.</span></span> <span data-ttu-id="59084-112">Office utilise alors l'interface utilisateur du complément COM au lieu de l'interface utilisateur de votre complément Office lorsque le complément COM est installé.</span><span class="sxs-lookup"><span data-stu-id="59084-112">Office will then use the COM add-in UI instead of your Office Add-in's UI when the COM add-in is installed.</span></span>

<span data-ttu-id="59084-113">L'exemple suivant montre comment spécifier un complément COM et un XLL comme équivalent.</span><span class="sxs-lookup"><span data-stu-id="59084-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="59084-114">Souvent, vous spécifierez à la fois de manière à ce que cet exemple montre les deux dans le contexte.</span><span class="sxs-lookup"><span data-stu-id="59084-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="59084-115">Ils sont identifiés par leur `ProgID` et `FileName` respectivement.</span><span class="sxs-lookup"><span data-stu-id="59084-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="59084-116">Pour plus d'informations sur la compatibilité des XLL, consultez [la rubrique faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l'utilisateur](../excel/make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="59084-116">For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="59084-117">Comportement équivalent pour les utilisateurs</span><span class="sxs-lookup"><span data-stu-id="59084-117">Equivalent behavior for users</span></span>

<span data-ttu-id="59084-118">Lorsqu'un complément COM équivalent est spécifié dans le manifeste du complément Office, Office supprime l'interface utilisateur de votre complément Office sur Windows lorsque le complément COM équivalent est installé.</span><span class="sxs-lookup"><span data-stu-id="59084-118">When an equivalent COM add-in is specified in the Office Add-in manifest, Office suppresses your Office Add-in's UI on Windows when the equivalent COM add-in is installed.</span></span> <span data-ttu-id="59084-119">Cela n'affecte pas l'interface utilisateur de votre complément Office sur d'autres plateformes, comme Online ou macOS.</span><span class="sxs-lookup"><span data-stu-id="59084-119">This does not affect your Office Add-in's UI on other platforms like online or macOS.</span></span> <span data-ttu-id="59084-120">Office masque uniquement les boutons du ruban et n'empêche pas l'installation.</span><span class="sxs-lookup"><span data-stu-id="59084-120">Office only hides the ribbon buttons and does not prevent installation.</span></span> <span data-ttu-id="59084-121">Par conséquent, votre complément Office continuera à apparaître aux emplacements d'IU suivants:</span><span class="sxs-lookup"><span data-stu-id="59084-121">Therefore your Office Add-in will still appear in the following UI locations:</span></span>

- <span data-ttu-id="59084-122">Sous **My Add-ins** car il est techniquement installé.</span><span class="sxs-lookup"><span data-stu-id="59084-122">Under **My add-ins** because it is technically installed.</span></span>
- <span data-ttu-id="59084-123">Comme entrée dans le gestionnaire de ruban.</span><span class="sxs-lookup"><span data-stu-id="59084-123">As an entry in the ribbon manager.</span></span>

<span data-ttu-id="59084-124">Les scénarios suivants décrivent ce qui se produit en fonction de la manière dont l'utilisateur acquiert le complément Office.</span><span class="sxs-lookup"><span data-stu-id="59084-124">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="59084-125">AppSource acquisition d'un complément Office</span><span class="sxs-lookup"><span data-stu-id="59084-125">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="59084-126">Si un utilisateur télécharge le complément Office à partir de AppSource, et que le complément COM équivalent est déjà installé, Office:</span><span class="sxs-lookup"><span data-stu-id="59084-126">If a user downloads the Office Add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="59084-127">Installez le complément Office.</span><span class="sxs-lookup"><span data-stu-id="59084-127">Install the Office Add-in.</span></span>
2. <span data-ttu-id="59084-128">Masquer l'interface utilisateur du complément Office dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="59084-128">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="59084-129">Afficher un appel pour l'utilisateur qui pointe vers le bouton du ruban de complément COM.</span><span class="sxs-lookup"><span data-stu-id="59084-129">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="59084-130">Déploiement centralisé du complément Office</span><span class="sxs-lookup"><span data-stu-id="59084-130">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="59084-131">Si un administrateur déploie le complément Office sur son client à l'aide d'un déploiement centralisé, et que le complément COM équivalent est déjà installé, l'utilisateur doit redémarrer Office pour qu'il voit les modifications.</span><span class="sxs-lookup"><span data-stu-id="59084-131">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes.</span></span> <span data-ttu-id="59084-132">Après le redémarrage d'Office, il peut:</span><span class="sxs-lookup"><span data-stu-id="59084-132">After Office restarts, it will:</span></span>

1. <span data-ttu-id="59084-133">Installez le complément Office.</span><span class="sxs-lookup"><span data-stu-id="59084-133">Install the Office Add-in.</span></span>
2. <span data-ttu-id="59084-134">Masquer l'interface utilisateur du complément Office dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="59084-134">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="59084-135">Afficher un appel pour l'utilisateur qui pointe vers le bouton du ruban de complément COM.</span><span class="sxs-lookup"><span data-stu-id="59084-135">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="59084-136">Document partagé avec un complément Office incorporé</span><span class="sxs-lookup"><span data-stu-id="59084-136">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="59084-137">Si un utilisateur a installé le complément COM, puis qu'il obtient un document partagé avec le complément Office incorporé, lorsqu'il ouvre le document, Office:</span><span class="sxs-lookup"><span data-stu-id="59084-137">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="59084-138">Inviter l'utilisateur à approuver le complément Office.</span><span class="sxs-lookup"><span data-stu-id="59084-138">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="59084-139">S'il est approuvé, le complément Office est installé.</span><span class="sxs-lookup"><span data-stu-id="59084-139">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="59084-140">Masquer l'interface utilisateur du complément Office dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="59084-140">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="59084-141">Autre comportement de complément COM</span><span class="sxs-lookup"><span data-stu-id="59084-141">Other COM add-in behavior</span></span>

<span data-ttu-id="59084-142">Si un utilisateur désinstalle le complément COM, office restaure l'interface utilisateur d'un complément Office sur Windows pour le complément Office installé équivalente.</span><span class="sxs-lookup"><span data-stu-id="59084-142">If a user uninstalls the COM add-in, then Office restores the Office Add-in UI on Windows for the equivalent installed Office Add-in.</span></span>

<span data-ttu-id="59084-143">Une fois que vous avez spécifié un complément COM équivalent pour votre complément Office, Office cesse de traiter les mises à jour pour votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="59084-143">Once you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="59084-144">L'utilisateur doit désinstaller l'ordre des compléments COM pour obtenir les dernières mises à jour pour le complément Office.</span><span class="sxs-lookup"><span data-stu-id="59084-144">The user must uninstall the COM add-in order to get the latest updates for the Office Add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="59084-145">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="59084-145">See also</span></span>

- [<span data-ttu-id="59084-146">Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l'utilisateur</span><span class="sxs-lookup"><span data-stu-id="59084-146">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
