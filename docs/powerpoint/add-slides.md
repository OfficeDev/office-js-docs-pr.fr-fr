---
title: Ajouter et supprimer des diapositives dans PowerPoint
description: Découvrez comment ajouter et supprimer des diapositives et spécifier le maître et la mise en page des nouvelles diapositives.
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: 9a8613997fc52ad6a30576b38c517a9c992f0e1b
ms.sourcegitcommit: ba4fb7087b9841d38bb46a99a63e88df49514a4d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/05/2021
ms.locfileid: "52779333"
---
# <a name="add-and-delete-slides-in-powerpoint"></a><span data-ttu-id="c16b3-103">Ajouter et supprimer des diapositives dans PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c16b3-103">Add and delete slides in PowerPoint</span></span>

<span data-ttu-id="c16b3-104">Un PowerPoint peut ajouter des diapositives à la présentation et éventuellement spécifier le maître des diapositives et la mise en page du maître utilisé pour la nouvelle diapositive.</span><span class="sxs-lookup"><span data-stu-id="c16b3-104">A PowerPoint add-in can add slides to the presentation and optionally specify which slide master, and which layout of the master, is used for the new slide.</span></span> <span data-ttu-id="c16b3-105">Le add-in peut également supprimer des diapositives.</span><span class="sxs-lookup"><span data-stu-id="c16b3-105">The add-in can also delete slides.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c16b3-106">Les API d’ajout de diapositives sont en [prévisualisation](../reference/requirement-sets/powerpoint-preview-apis.md) et ne sont pas disponibles pour les modules de production. L’API *de suppression des* diapositives a été publiée.</span><span class="sxs-lookup"><span data-stu-id="c16b3-106">The APIs for adding slides are in [preview](../reference/requirement-sets/powerpoint-preview-apis.md) and not available for production add-ins. The API for *deleting* slides has been released.</span></span>

<span data-ttu-id="c16b3-107">Les API d’ajout de diapositives sont principalement utilisées dans les scénarios où les ID des formes de base et des mises en page des diapositives de la présentation sont connus au moment du codage ou se trouvent dans une source de données lors de l’runtime.</span><span class="sxs-lookup"><span data-stu-id="c16b3-107">The APIs for adding slides are primarily used in scenarios where the IDs of the slide masters and layouts in the presentation are known at coding time or can be found in a data source at runtime.</span></span> <span data-ttu-id="c16b3-108">Dans ce cas, vous ou le client devez créer et gérer une source de données qui met en corrélation le critère de sélection (par exemple, les noms ou les images des formes de base et des mises en page des diapositives) avec les ID des formes de base et des mises en page des diapositives.</span><span class="sxs-lookup"><span data-stu-id="c16b3-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as the names or images of slide masters and layouts) with the IDs of the slide masters and layouts.</span></span> <span data-ttu-id="c16b3-109">Les API peuvent également être utilisées dans les scénarios où l’utilisateur peut insérer des diapositives qui utilisent le maître des diapositives par défaut et la mise en page par défaut du maître, et dans les scénarios où l’utilisateur peut sélectionner une diapositive existante et en créer une nouvelle avec le même maître et la même mise en page de diapositives (mais pas le même contenu).</span><span class="sxs-lookup"><span data-stu-id="c16b3-109">The APIs can also be used in scenarios where the user can insert slides that use the default slide master and the master's default layout, and in scenarios where the user can select an existing slide and create a new one with the same slide master and layout (but not the same content).</span></span> <span data-ttu-id="c16b3-110">Pour [plus d’informations à](#selecting-which-slide-master-and-layout-to-use) ce sujet, voir Sélection du maître des diapositives et de la mise en page à utiliser.</span><span class="sxs-lookup"><span data-stu-id="c16b3-110">See [Selecting which slide master and layout to use](#selecting-which-slide-master-and-layout-to-use) for more information about this.</span></span>

## <a name="add-a-slide-with-slidecollectionadd-preview"></a><span data-ttu-id="c16b3-111">Ajouter une diapositive avec SlideCollection.add (aperçu)</span><span class="sxs-lookup"><span data-stu-id="c16b3-111">Add a slide with SlideCollection.add (preview)</span></span>

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

<span data-ttu-id="c16b3-112">Ajoutez des diapositives avec [la méthode SlideCollection.add.](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)</span><span class="sxs-lookup"><span data-stu-id="c16b3-112">Add slides with the [SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) method.</span></span> <span data-ttu-id="c16b3-113">Voici un exemple simple dans lequel une diapositive qui utilise le maître des diapositives par défaut de la présentation et la première mise en page de ce maître est ajoutée.</span><span class="sxs-lookup"><span data-stu-id="c16b3-113">The following is a simple example in which a slide that uses the presentation's default slide master and the first layout of that master is added.</span></span> <span data-ttu-id="c16b3-114">La méthode ajoute toujours de nouvelles diapositives à la fin de la présentation.</span><span class="sxs-lookup"><span data-stu-id="c16b3-114">The method always adds new slides to the end of the presentation.</span></span> <span data-ttu-id="c16b3-115">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="c16b3-115">The following is an example:</span></span>

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="selecting-which-slide-master-and-layout-to-use"></a><span data-ttu-id="c16b3-116">Sélection du maître des diapositives et de la mise en page à utiliser</span><span class="sxs-lookup"><span data-stu-id="c16b3-116">Selecting which slide master and layout to use</span></span>

<span data-ttu-id="c16b3-117">Utilisez le [paramètre AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) pour contrôler le maître des diapositives qui est utilisé pour la nouvelle diapositive et la mise en page dans le master.</span><span class="sxs-lookup"><span data-stu-id="c16b3-117">Use the [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) parameter to control which slide master is used for the new slide and which layout within the master is used.</span></span> <span data-ttu-id="c16b3-118">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="c16b3-118">The following is an example.</span></span> <span data-ttu-id="c16b3-119">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="c16b3-119">Note the following about this code:</span></span>

- <span data-ttu-id="c16b3-120">Vous pouvez inclure l’une ou l’autre des propriétés de l’objet ou les `AddSlideOptions` deux.</span><span class="sxs-lookup"><span data-stu-id="c16b3-120">You can include either or both the properties of the `AddSlideOptions` object.</span></span>
- <span data-ttu-id="c16b3-121">Si les deux propriétés sont utilisées, la disposition spécifiée doit appartenir à la forme de base spécifiée ou une erreur est lancée.</span><span class="sxs-lookup"><span data-stu-id="c16b3-121">If both properties are used, then the specified layout must belong to the specified master or an error is thrown.</span></span>
- <span data-ttu-id="c16b3-122">Si la propriété n’est pas présente (ou si sa valeur est une chaîne vide), le curseur de diapositive par défaut est utilisé et doit être une mise en page de `masterId` `layoutId` ce dernier.</span><span class="sxs-lookup"><span data-stu-id="c16b3-122">If the `masterId` property isn't present (or its value is an empty string), then the default slide master is used and the `layoutId` must be a layout of that slide master.</span></span>
- <span data-ttu-id="c16b3-123">Le maître des diapositives par défaut est celui utilisé par la dernière diapositive de la présentation.</span><span class="sxs-lookup"><span data-stu-id="c16b3-123">The default slide master is the slide master used by the last slide in the presentation.</span></span> <span data-ttu-id="c16b3-124">(Dans le cas rare où il n’y a actuellement aucune diapositive dans la présentation, le maître des diapositives par défaut est le premier maître des diapositives de la présentation.)</span><span class="sxs-lookup"><span data-stu-id="c16b3-124">(In the unusual case where there are currently no slides in the presentation, then the default slide master is the first slide master in the presentation.)</span></span>
- <span data-ttu-id="c16b3-125">Si la propriété n’est pas présente (ou si sa valeur est une chaîne vide), la première disposition de la forme de base spécifiée par la forme de base `layoutId` `masterId` est utilisée.</span><span class="sxs-lookup"><span data-stu-id="c16b3-125">If the `layoutId` property isn't present (or its value is an empty string), then the first layout of the master that is specified by the `masterId` is used.</span></span>
- <span data-ttu-id="c16b3-126">Les deux propriétés sont des chaînes de l’une des trois formes possibles : ***nnnnnnnnnn\*#\*\*, \* *#* mmmmmmmmmmm*** ou \**_nnnnnnnnnn_ #* mmmmmmmmm\*\*\*, où *nnnnnnnnnn* est l’ID de la forme de base ou de la disposition (généralement 10 chiffres) et *mmmmmmmmmmm* est l’ID de création de la forme de base ou de la disposition (généralement 6 à 10 chiffres).</span><span class="sxs-lookup"><span data-stu-id="c16b3-126">Both properties are strings of one of three possible forms: \***nnnnnnnnnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnnnnnnnnn_#* mmmmmmmmm\*\*\*, where *nnnnnnnnnn* is the master's or layout's ID (typically 10 digits) and *mmmmmmmmm* is the master's or layout's creation ID (typically 6 - 10 digits).</span></span> <span data-ttu-id="c16b3-127">Voici quelques exemples `2147483690#2908289500` : `2147483690#` , et `#2908289500` .</span><span class="sxs-lookup"><span data-stu-id="c16b3-127">Some examples are `2147483690#2908289500`, `2147483690#`, and `#2908289500`.</span></span>

```javascript
async function addSlide() {
    await PowerPoint.run(async function(context) {
        context.presentation.slides.add({
            slideMasterId: "2147483690#2908289500",
            layoutId: "2147483691#2499880"
        });
    
        await context.sync();
    });
}
```

<span data-ttu-id="c16b3-128">Il n’existe aucun moyen pratique pour les utilisateurs de découvrir l’ID ou l’ID de création d’un curseur de diapositive ou d’une mise en page.</span><span class="sxs-lookup"><span data-stu-id="c16b3-128">There is no practical way that users can discover the ID or creation ID of a slide master or layout.</span></span> <span data-ttu-id="c16b3-129">Pour cette raison, vous ne pouvez utiliser le paramètre que lorsque vous connaissez les ID au moment du codage ou que votre application peut les découvrir au moment de `AddSlideOptions` l’utilisation.</span><span class="sxs-lookup"><span data-stu-id="c16b3-129">For this reason, you can really only use the `AddSlideOptions` parameter when either you know the IDs at coding time or your add-in can discover them at runtime.</span></span> <span data-ttu-id="c16b3-130">Étant donné que les utilisateurs ne sont pas censés mémoriser les ID, vous avez également besoin d’un moyen pour permettre à l’utilisateur de sélectionner des diapositives, par exemple par son nom ou par une image, puis de corréler chaque titre ou image avec l’ID de la diapositive.</span><span class="sxs-lookup"><span data-stu-id="c16b3-130">Because users can't be expected to memorize the IDs, you also need a way to enable the user to select slides, perhaps by name or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="c16b3-131">Par conséquent, le paramètre est principalement utilisé dans les scénarios dans lesquels le module est conçu pour fonctionner avec un ensemble spécifique de formes de base et de mises en page dont les ID sont `AddSlideOptions` connus.</span><span class="sxs-lookup"><span data-stu-id="c16b3-131">Accordingly, the `AddSlideOptions` parameter is primarily used in scenarios in which the add-in is designed to work with a specific set of slide masters and layouts whose IDs are known.</span></span> <span data-ttu-id="c16b3-132">Dans ce cas, vous ou le client devez créer et gérer une source de données qui met en corrélation un critère de sélection (tel que le maître des diapositives et les noms ou images de mise en page) avec les ID ou les ID de création correspondants.</span><span class="sxs-lookup"><span data-stu-id="c16b3-132">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as slide master and layout names or images) with the corresponding IDs or creation IDs.</span></span>

#### <a name="have-the-user-choose-a-matching-slide"></a><span data-ttu-id="c16b3-133">Faire en cas de choix d’une diapositive correspondante par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="c16b3-133">Have the user choose a matching slide</span></span>

<span data-ttu-id="c16b3-134">Si votre add-in peut être utilisé dans des scénarios où la nouvelle diapositive doit  utiliser la même combinaison de formes de base et de mise en page que celle utilisée par une diapositive existante, votre add-in peut (1) invite l’utilisateur à sélectionner une diapositive et (2) lit les ID du maître et de la mise en page des diapositives.</span><span class="sxs-lookup"><span data-stu-id="c16b3-134">If your add-in can be used in scenarios where the new slide should use the same combination of slide master and layout that is used by an *existing* slide, then your add-in can (1) prompt the user to select a slide and (2) read the IDs of the slide master and layout.</span></span> <span data-ttu-id="c16b3-135">Les étapes suivantes montrent comment lire les ID et ajouter une diapositive avec une forme de base et une mise en page correspondantes.</span><span class="sxs-lookup"><span data-stu-id="c16b3-135">The following steps show how to read the IDs and add a slide with a matching master and layout.</span></span>

1. <span data-ttu-id="c16b3-136">Créez une méthode pour obtenir l’index de la diapositive sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="c16b3-136">Create a method to get the index of the selected slide.</span></span> <span data-ttu-id="c16b3-137">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="c16b3-137">The following is an example.</span></span> <span data-ttu-id="c16b3-138">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="c16b3-138">Note about this code:</span></span>

    - <span data-ttu-id="c16b3-139">Il utilise la [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) des API JavaScript communes.</span><span class="sxs-lookup"><span data-stu-id="c16b3-139">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="c16b3-140">L’appel `getSelectedDataAsync` est incorporé dans une fonction de renvoi de promesse.</span><span class="sxs-lookup"><span data-stu-id="c16b3-140">The call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="c16b3-141">Pour plus d’informations sur la raison et la façon de le faire, voir [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span><span class="sxs-lookup"><span data-stu-id="c16b3-141">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="c16b3-142">`getSelectedDataAsync` renvoie un tableau car plusieurs diapositives peuvent être sélectionnées.</span><span class="sxs-lookup"><span data-stu-id="c16b3-142">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="c16b3-143">Dans ce scénario, l’utilisateur n’en a sélectionné qu’une seule, de sorte que le code obtient la première (0e) diapositive, qui est la seule sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="c16b3-143">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="c16b3-144">La valeur de la diapositive est la valeur 1 que l’utilisateur voit en regard de la diapositive dans le volet de `index` miniatures.</span><span class="sxs-lookup"><span data-stu-id="c16b3-144">The `index` value of the slide is the 1-based value the user sees beside the slide in the thumbnails pane.</span></span>

    ```javascript
    function getSelectedSlideIndex() {
        return new OfficeExtension.Promise<number>(function(resolve, reject) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function(asyncResult) {
                try {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(console.error(asyncResult.error.message));
                    } else {
                        resolve(asyncResult.value.slides[0].index);
                    }
                } 
                catch (error) {
                    reject(console.log(error));
                }
            });
        });
    }
    ```

2. <span data-ttu-id="c16b3-145">Appelez votre nouvelle fonction à [l’PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) de la fonction principale qui ajoute la diapositive.</span><span class="sxs-lookup"><span data-stu-id="c16b3-145">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function that adds the slide.</span></span> <span data-ttu-id="c16b3-146">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="c16b3-146">The following is an example:</span></span>

    ```javascript
    async function addSlideWithMatchingLayout() {
        await PowerPoint.run(async function(context) {
    
            let selectedSlideIndex = await getSelectedSlideIndex();
        
            // Decrement the index because the value returned by getSelectedSlideIndex()
            // is 1-based, but SlideCollection.getItemAt() is 0-based.
            const realSlideIndex = selectedSlideIndex - 1;
            const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex).load("slideMaster/id, layout/id");
        
            await context.sync();
        
            context.presentation.slides.add({
                slideMasterId: selectedSlide.slideMaster.id,
                layoutId: selectedSlide.layout.id
            });
        
            await context.sync();
        });
    }
    ```

## <a name="delete-slides"></a><span data-ttu-id="c16b3-147">Supprimer des diapositives</span><span class="sxs-lookup"><span data-stu-id="c16b3-147">Delete slides</span></span>

<span data-ttu-id="c16b3-148">Supprimez une diapositive en obtenant une référence à l’objet [Slide](/javascript/api/powerpoint/powerpoint.slide) qui représente la diapositive et appelez la `Slide.delete` méthode.</span><span class="sxs-lookup"><span data-stu-id="c16b3-148">Delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="c16b3-149">Voici un exemple dans lequel la quatrième diapositive est supprimée :</span><span class="sxs-lookup"><span data-stu-id="c16b3-149">The following is an example in which the 4th slide is deleted:</span></span>

```javascript
async function deleteSlide() {
    await PowerPoint.run(async function(context) {

        // The slide index is zero-based. 
        const slide = context.presentation.slides.getItemAt(3);
        slide.delete();

        await context.sync();
    });
}
```
