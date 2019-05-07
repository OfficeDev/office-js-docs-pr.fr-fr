---
title: Faire en sorte que votre complément Excel soit compatible avec un complément COM existant
description: Activer la compatibilité avec un complément COM équivalent doté de la même fonctionnalité que votre complément Excel
ms.date: 05/06/2019
localization_priority: Normal
ms.openlocfilehash: 0890e14466a2cd8f5aff2d1bcf307a43cff28127
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628171"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a><span data-ttu-id="1e9ed-103">Faire en sorte que votre complément Office soit compatible avec un complément COM existant (aperçu)</span><span class="sxs-lookup"><span data-stu-id="1e9ed-103">Make your Office Add-in compatible with an existing COM add-in (preview)</span></span>

<span data-ttu-id="1e9ed-104">Si vous disposez d’un complément COM existant, vous pouvez créer une fonctionnalité équivalente dans votre complément Excel afin d’étendre les fonctionnalités de votre solution à d’autres plateformes, comme Online ou macOS.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-104">If you have an existing COM add-in, you can build equivalent functionality in your Excel add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="1e9ed-105">Toutefois, les compléments Excel ne disposent pas de toutes les fonctionnalités disponibles dans les compléments COM. Votre complément COM peut fournir une meilleure expérience que le complément Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-105">However, Excel add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Excel add-in on Windows.</span></span>

<span data-ttu-id="1e9ed-106">Vous pouvez configurer votre complément Excel de sorte que, lorsqu’un complément COM équivalent est déjà installé sur l’ordinateur de l’utilisateur, Office exécute le complément COM au lieu de votre complément Excel.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-106">You can configure your Excel add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Excel add-in.</span></span> <span data-ttu-id="1e9ed-107">Le complément COM est appelé «équivalent», car Office effectuera une transition transparente entre le complément COM et le complément Excel en fonction de ce qui est installé sur Windows.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-107">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Excel add-in depending on which is installed on Windows.</span></span>

[!include[COM add-in and XLL UDF compatibility requirements note](../includes/xll-compatibility-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="1e9ed-108">Spécifier un complément COM équivalent dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="1e9ed-108">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="1e9ed-109">Pour activer la compatibilité avec un complément COM existant, identifiez le complément COM équivalent dans le manifeste de votre complément Excel.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-109">To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Excel add-in.</span></span> <span data-ttu-id="1e9ed-110">Office utilise ensuite le complément COM au lieu de votre complément Excel lors de l’exécution de Windows.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-110">Then Office will use the COM add-in instead of your Excel add-in when running on Windows.</span></span>

<span data-ttu-id="1e9ed-111">Spécifiez `ProgID` le du complément COM équivalent.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-111">Specify the `ProgID` of the equivalent COM add-in.</span></span> <span data-ttu-id="1e9ed-112">Office utilise ensuite l’interface utilisateur du complément COM au lieu de l’interface utilisateur de votre complément Excel lorsque le complément COM est installé.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-112">Office will then use the COM add-in UI instead of your Excel add-in's UI when the COM add-in is installed.</span></span>

<span data-ttu-id="1e9ed-113">L’exemple suivant montre comment spécifier un complément COM et un XLL comme équivalent.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="1e9ed-114">Souvent, vous spécifierez à la fois de manière à ce que cet exemple montre les deux dans le contexte.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="1e9ed-115">Ils sont identifiés par leur `ProgID` et `FileName` respectivement.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="1e9ed-116">Pour plus d’informations sur la compatibilité des XLL, consultez [la rubrique faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l’utilisateur](../excel/make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="1e9ed-116">For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

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

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="1e9ed-117">Comportement équivalent pour les utilisateurs</span><span class="sxs-lookup"><span data-stu-id="1e9ed-117">Equivalent behavior for users</span></span>

<span data-ttu-id="1e9ed-118">Lorsqu’un complément COM équivalent est spécifié dans le manifeste de complément Excel, Office supprime l’interface utilisateur de votre complément Excel sur Windows lorsque le complément COM équivalent est installé.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-118">When an equivalent COM add-in is specified in the Excel add-in manifest, Office suppresses your Excel add-in's UI on Windows when the equivalent COM add-in is installed.</span></span> <span data-ttu-id="1e9ed-119">Cela n’affecte pas l’interface utilisateur de votre complément Excel sur d’autres plateformes, comme Online ou macOS.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-119">This does not affect your Excel add-in's UI on other platforms like online or macOS.</span></span> <span data-ttu-id="1e9ed-120">Office masque uniquement les boutons du ruban et n’empêche pas l’installation.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-120">Office only hides the ribbon buttons and does not prevent installation.</span></span> <span data-ttu-id="1e9ed-121">Par conséquent, votre complément Excel apparaîtra toujours dans les emplacements d’IU suivants:</span><span class="sxs-lookup"><span data-stu-id="1e9ed-121">Therefore your Excel add-in will still appear in the following UI locations:</span></span>

- <span data-ttu-id="1e9ed-122">Sous **My Add-ins** car il est techniquement installé.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-122">Under **My add-ins** because it is technically installed.</span></span>
- <span data-ttu-id="1e9ed-123">Comme entrée dans le gestionnaire de ruban.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-123">As an entry in the ribbon manager.</span></span>

<span data-ttu-id="1e9ed-124">Les scénarios suivants décrivent ce qui se produit en fonction de la manière dont l’utilisateur acquiert le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-124">The following scenarios describe what happens depending on how the user acquires the Excel add-in.</span></span>

### <a name="appsource-acquisition-of-an-excel-add-in"></a><span data-ttu-id="1e9ed-125">AppSource acquisition d’un complément Excel</span><span class="sxs-lookup"><span data-stu-id="1e9ed-125">AppSource acquisition of an Excel add-in</span></span>

<span data-ttu-id="1e9ed-126">Si un utilisateur télécharge le complément Excel à partir de AppSource, et que le complément COM équivalent est déjà installé, Office:</span><span class="sxs-lookup"><span data-stu-id="1e9ed-126">If a user downloads the Excel add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="1e9ed-127">Installez le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-127">Install the Excel add-in.</span></span>
2. <span data-ttu-id="1e9ed-128">Masquer l’interface utilisateur du complément Excel dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-128">Hide the Excel add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="1e9ed-129">Afficher un appel pour l’utilisateur qui pointe vers le bouton du ruban de complément COM.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-129">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-excel-add-in"></a><span data-ttu-id="1e9ed-130">Déploiement centralisé d’un complément Excel</span><span class="sxs-lookup"><span data-stu-id="1e9ed-130">Centralized deployment of Excel add-in</span></span>

<span data-ttu-id="1e9ed-131">Si un administrateur déploie le complément Excel sur son client à l’aide d’un déploiement centralisé, et que le complément COM équivalent est déjà installé, l’utilisateur doit redémarrer Office pour qu’il voit les modifications.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-131">If an admin deploys the Excel add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes.</span></span> <span data-ttu-id="1e9ed-132">Après le redémarrage d’Office, il peut:</span><span class="sxs-lookup"><span data-stu-id="1e9ed-132">After Office restarts, it will:</span></span>

1. <span data-ttu-id="1e9ed-133">Installez le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-133">Install the Excel add-in.</span></span>
2. <span data-ttu-id="1e9ed-134">Masquer l’interface utilisateur du complément Excel dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-134">Hide the Excel add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="1e9ed-135">Afficher un appel pour l’utilisateur qui pointe vers le bouton du ruban de complément COM.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-135">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-excel-add-in"></a><span data-ttu-id="1e9ed-136">Document partagé avec un complément Excel incorporé</span><span class="sxs-lookup"><span data-stu-id="1e9ed-136">Document shared with embedded Excel add-in</span></span>

<span data-ttu-id="1e9ed-137">Si un utilisateur a installé le complément COM, puis qu’il obtient un document partagé avec le complément Excel incorporé, lorsqu’il ouvre le document, Office:</span><span class="sxs-lookup"><span data-stu-id="1e9ed-137">If a user has the COM add-in installed, and then gets a shared document with the embedded Excel add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="1e9ed-138">Inviter l’utilisateur à approuver le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-138">Prompt the user to trust the Excel add-in.</span></span>
2. <span data-ttu-id="1e9ed-139">S’il est approuvé, le complément Excel s’installe.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-139">If trusted, the Excel add-in will install.</span></span>
3. <span data-ttu-id="1e9ed-140">Masquer l’interface utilisateur du complément Excel dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-140">Hide the Excel add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="1e9ed-141">Autre comportement de complément COM</span><span class="sxs-lookup"><span data-stu-id="1e9ed-141">Other COM add-in behavior</span></span>

<span data-ttu-id="1e9ed-142">Si un utilisateur désinstalle le complément COM, office restaure l’interface utilisateur d’un complément Excel sur Windows pour le complément Excel installé équivalente.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-142">If a user uninstalls the COM add-in, then Office restores the Excel add-in UI on Windows for the equivalent installed Excel add-in.</span></span>

<span data-ttu-id="1e9ed-143">Une fois que vous avez spécifié un complément COM équivalent pour votre complément Excel, Office cesse de traiter les mises à jour pour votre complément Excel.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-143">Once you specify an equivalent COM add-in for your Excel add-in, Office stops processing updates for your Excel add-in.</span></span> <span data-ttu-id="1e9ed-144">L’utilisateur doit désinstaller l’ordre des compléments COM pour obtenir les dernières mises à jour pour le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="1e9ed-144">The user must uninstall the COM add-in order to get the latest updates for the Excel add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="1e9ed-145">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1e9ed-145">See also</span></span>

- [<span data-ttu-id="1e9ed-146">Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="1e9ed-146">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
