---
title: Utiliser des balises personnalisées sur les présentations, diapositives et formes dans PowerPoint
description: Découvrez comment utiliser des balises pour des métadonnées personnalisées sur les présentations, les diapositives et les formes.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fbb13e67da1f7962fc2c0b8d45689f259b015014
ms.sourcegitcommit: 58d394fa49308ecf93cd53f7d3fb6e316ff56209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/16/2021
ms.locfileid: "51876857"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a><span data-ttu-id="66ad8-103">Utiliser des balises personnalisées pour les présentations, les diapositives et les formes dans PowerPoint</span><span class="sxs-lookup"><span data-stu-id="66ad8-103">Use custom tags for presentations, slides, and shapes in PowerPoint</span></span>

<span data-ttu-id="66ad8-104">Un add-in peut joindre des métadonnées personnalisées, sous la forme de paires clé-valeur, appelées « balises », à des présentations, des diapositives spécifiques et des formes spécifiques sur une diapositive.</span><span class="sxs-lookup"><span data-stu-id="66ad8-104">An add-in can attach custom metadata, in the form of key-value pairs, called "tags", to presentations, specific slides, and specific shapes on a slide.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="66ad8-105">Les API pour les balises sont en prévisualisation.</span><span class="sxs-lookup"><span data-stu-id="66ad8-105">The APIs for tags are in preview.</span></span> <span data-ttu-id="66ad8-106">Testez-les dans un environnement de développement ou de test, mais ne les ajoutez pas à un module de production.</span><span class="sxs-lookup"><span data-stu-id="66ad8-106">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>

<span data-ttu-id="66ad8-107">Il existe deux scénarios principaux pour l'utilisation de balises :</span><span class="sxs-lookup"><span data-stu-id="66ad8-107">There are two main scenarios for using tags:</span></span>

- <span data-ttu-id="66ad8-108">Lorsqu'elle est appliquée à une diapositive ou à une forme, une balise permet de classer l'objet pour le traitement par lots.</span><span class="sxs-lookup"><span data-stu-id="66ad8-108">When applied to a slide or a shape, a tag enables the object to be categorized for batch processing.</span></span> <span data-ttu-id="66ad8-109">Par exemple, supposons qu'une présentation possède des diapositives qui doivent être incluses dans les présentations de la région Est, mais pas de la région Ouest.</span><span class="sxs-lookup"><span data-stu-id="66ad8-109">For example, suppose a presentation has some slides that should be included in presentations to the East region but not the West region.</span></span> <span data-ttu-id="66ad8-110">De même, il existe d'autres diapositives qui doivent être affichées uniquement à l'Ouest.</span><span class="sxs-lookup"><span data-stu-id="66ad8-110">Similarly, there are alternative slides that should be shown only to the West.</span></span> <span data-ttu-id="66ad8-111">Votre application peut créer une balise avec la clé et la valeur et l'appliquer aux diapositives qui ne doivent être utilisées `REGION` `East` qu'à l'Est.</span><span class="sxs-lookup"><span data-stu-id="66ad8-111">Your add-in can create a tag with the key `REGION` and the value `East` and apply it to the slides that should only be used in the East.</span></span> <span data-ttu-id="66ad8-112">La valeur de la balise est définie pour les diapositives qui doivent uniquement être `West` affichées dans la région Ouest.</span><span class="sxs-lookup"><span data-stu-id="66ad8-112">The tag's value is set to `West` for the slides that should only be shown to the West region.</span></span> <span data-ttu-id="66ad8-113">Juste avant une présentation à l'Est, un bouton du add-in exécute un code qui pare toutes les diapositives en vérifiant la valeur de la `REGION` balise.</span><span class="sxs-lookup"><span data-stu-id="66ad8-113">Just before a presentation to the East, a button in the add-in runs code that loops through all the slides checking the value of the `REGION` tag.</span></span> <span data-ttu-id="66ad8-114">Diapositives dans laquelle la région `West` est supprimée.</span><span class="sxs-lookup"><span data-stu-id="66ad8-114">Slides where the region is `West` are deleted.</span></span> <span data-ttu-id="66ad8-115">L'utilisateur ferme ensuite le module et démarre le diaporama.</span><span class="sxs-lookup"><span data-stu-id="66ad8-115">The user then closes the add-in and starts the slide show.</span></span>
- <span data-ttu-id="66ad8-116">Lorsqu'elle est appliquée à une présentation, une balise est en fait une propriété personnalisée dans le document de présentation (semblable à [une propriété](/javascript/api/word/word.customproperty) personnalisée dans Word).</span><span class="sxs-lookup"><span data-stu-id="66ad8-116">When applied to a presentation, a tag is effectively a custom property in the presentation document (similar to a [CustomProperty](/javascript/api/word/word.customproperty) in Word).</span></span>

## <a name="tag-slides-and-shapes"></a><span data-ttu-id="66ad8-117">Baliser les diapositives et les formes</span><span class="sxs-lookup"><span data-stu-id="66ad8-117">Tag slides and shapes</span></span>

<span data-ttu-id="66ad8-118">Une balise est une paire clé-valeur, où la valeur est toujours de type et est représentée `string` par un [objet Tag.](/javascript/api/powerpoint/powerpoint.tag)</span><span class="sxs-lookup"><span data-stu-id="66ad8-118">A tag is a key-value pair, where the value is always of type `string` and is represented by a [Tag](/javascript/api/powerpoint/powerpoint.tag) object.</span></span> <span data-ttu-id="66ad8-119">Chaque type d'objet parent, tel qu'un objet [Presentation,](/javascript/api/powerpoint/powerpoint.presentation) [Slide](/javascript/api/powerpoint/powerpoint.slide)ou [Shape,](/javascript/api/powerpoint/powerpoint.shape) possède une propriété `tags` de type [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).</span><span class="sxs-lookup"><span data-stu-id="66ad8-119">Each type of parent object, such as a [Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide), or [Shape](/javascript/api/powerpoint/powerpoint.shape) object, has a `tags` property of type [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).</span></span>

### <a name="add-update-and-delete-tags"></a><span data-ttu-id="66ad8-120">Ajouter, mettre à jour et supprimer des balises</span><span class="sxs-lookup"><span data-stu-id="66ad8-120">Add, update, and delete tags</span></span>

<span data-ttu-id="66ad8-121">Pour ajouter une balise à un objet, appelez la méthode [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) de la propriété de l'objet `tags` parent.</span><span class="sxs-lookup"><span data-stu-id="66ad8-121">To add a tag to an object, call the [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) method of the parent object's `tags` property.</span></span> <span data-ttu-id="66ad8-122">Le code suivant ajoute deux balises à la première diapositive d'une présentation.</span><span class="sxs-lookup"><span data-stu-id="66ad8-122">The following code adds two tags to the first slide of a presentation.</span></span> <span data-ttu-id="66ad8-123">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="66ad8-123">About this code, note:</span></span>

- <span data-ttu-id="66ad8-124">Le premier paramètre de la méthode est la clé de la paire `add` clé-valeur.</span><span class="sxs-lookup"><span data-stu-id="66ad8-124">The first parameter of the `add` method is the key in the key-value pair.</span></span> 
- <span data-ttu-id="66ad8-125">Le deuxième paramètre est la valeur.</span><span class="sxs-lookup"><span data-stu-id="66ad8-125">The second parameter is the value.</span></span>
- <span data-ttu-id="66ad8-126">La clé est en lettres majuscules.</span><span class="sxs-lookup"><span data-stu-id="66ad8-126">The key is in uppercase letters.</span></span> <span data-ttu-id="66ad8-127">Cela n'est pas strictement obligatoire pour la méthode ; toutefois, la clé est toujours stockée par PowerPoint en tant que minuscules, et certaines méthodes liées aux balises nécessitent que la clé soit exprimée en minuscules . Nous vous recommandons donc, en tant que meilleure pratique, d'utiliser toujours des minuscules dans votre code pour une clé de `add` balise. </span><span class="sxs-lookup"><span data-stu-id="66ad8-127">This isn't strictly mandatory for the `add` method; however, the key is always stored by PowerPoint as uppercase, and *some tag-related methods do require that the key be expressed in uppercase*, so we recommend as a best practice that you always use uppercase in your code for a tag key.</span></span>

```javascript
async function addMultipleSlideTags() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("OCEAN", "Arctic");
    slide.tags.add("PLANET", "Jupiter");

    await context.sync();
  });
}
```

<span data-ttu-id="66ad8-128">La `add` méthode est également utilisée pour mettre à jour une balise.</span><span class="sxs-lookup"><span data-stu-id="66ad8-128">The `add` method is also used to update a tag.</span></span> <span data-ttu-id="66ad8-129">Le code suivant modifie la valeur de la `PLANET` balise.</span><span class="sxs-lookup"><span data-stu-id="66ad8-129">The following code changes the value of the `PLANET` tag.</span></span>

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

<span data-ttu-id="66ad8-130">Pour supprimer une balise, appelez la méthode sur son objet parent et passez la clé de la `delete` balise en tant que `TagsCollection` paramètre.</span><span class="sxs-lookup"><span data-stu-id="66ad8-130">To delete a tag, call the `delete` method on it's parent `TagsCollection` object and pass the key of the tag as the parameter.</span></span> <span data-ttu-id="66ad8-131">Pour obtenir un exemple, voir [Définir des métadonnées personnalisées sur la présentation.](#set-custom-metadata-on-the-presentation)</span><span class="sxs-lookup"><span data-stu-id="66ad8-131">For an example, see [Set custom metadata on the presentation](#set-custom-metadata-on-the-presentation).</span></span>

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a><span data-ttu-id="66ad8-132">Utiliser des balises pour traiter de manière sélective les diapositives et les formes</span><span class="sxs-lookup"><span data-stu-id="66ad8-132">Use tags to selectively process slides and shapes</span></span>

<span data-ttu-id="66ad8-133">Envisagez le scénario suivant : Contoso Consulting présente une présentation qu'il présente à tous les nouveaux clients.</span><span class="sxs-lookup"><span data-stu-id="66ad8-133">Consider the following scenario: Contoso Consulting has a presentation they show to all new customers.</span></span> <span data-ttu-id="66ad8-134">Toutefois, certaines diapositives ne doivent être affichées qu'aux clients qui ont payé l'état « premium ».</span><span class="sxs-lookup"><span data-stu-id="66ad8-134">But some slides should only be shown to customers that have paid for "premium" status.</span></span> <span data-ttu-id="66ad8-135">Avant d'afficher la présentation aux clients non premium, ils en font une copie et suppriment les diapositives que seuls les clients premium doivent voir.</span><span class="sxs-lookup"><span data-stu-id="66ad8-135">Before showing the presentation to non-premium customers, they make a copy of it and delete the slides that only premium customers should see.</span></span> <span data-ttu-id="66ad8-136">Un add-in permet à Contoso de baliser les diapositives qui sont pour les clients premium et de supprimer ces diapositives si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="66ad8-136">An add-in enables Contoso to tag which slides are for premium customers and to delete these slides when needed.</span></span> <span data-ttu-id="66ad8-137">La liste suivante décrit les principales étapes de codage pour créer cette fonctionnalité.</span><span class="sxs-lookup"><span data-stu-id="66ad8-137">The following list outlines the major coding steps to create this functionality.</span></span>

1. <span data-ttu-id="66ad8-138">Créez une méthode qui balise la diapositive actuellement sélectionnée comme prévu pour les `Premium` clients.</span><span class="sxs-lookup"><span data-stu-id="66ad8-138">Create a method that tags the currently selected slide as intended for `Premium` customers.</span></span> <span data-ttu-id="66ad8-139">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="66ad8-139">About this code, note:</span></span>

    - <span data-ttu-id="66ad8-140">La `getSelectedSlideIndex` fonction est définie à l'étape suivante.</span><span class="sxs-lookup"><span data-stu-id="66ad8-140">The `getSelectedSlideIndex` function is defined in the next step.</span></span> <span data-ttu-id="66ad8-141">Elle renvoie l'index de base 1 de la diapositive actuellement sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="66ad8-141">It returns the 1-based index of the currently selected slide.</span></span>
    - <span data-ttu-id="66ad8-142">La valeur renvoyée par la fonction doit être décrémentée car la méthode `getSelectedSlideIndex` [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) est basée sur 0.</span><span class="sxs-lookup"><span data-stu-id="66ad8-142">The value returned by the `getSelectedSlideIndex` function has to be decremented because the [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) method is 0-based.</span></span>

    ```javascript
    async function addTagToSelectedSlide() {
      await PowerPoint.run(async function(context) {
        let selectedSlideIndex = await getSelectedSlideIndex();
        selectedSlideIndex = selectedSlideIndex - 1;
        const slide = context.presentation.slides.getItemAt(selectedSlideIndex);
        slide.tags.add("CUSTOMER_TYPE", "Premium");
    
        await context.sync();
      });
    }
    ```

2. <span data-ttu-id="66ad8-143">Le code suivant crée une méthode pour obtenir l'index de la diapositive sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="66ad8-143">The following code creates a method to get the index of the selected slide.</span></span> <span data-ttu-id="66ad8-144">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="66ad8-144">About this code, note:</span></span>

    - <span data-ttu-id="66ad8-145">Il utilise la [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) des API JavaScript communes.</span><span class="sxs-lookup"><span data-stu-id="66ad8-145">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="66ad8-146">L'appel `getSelectedDataAsync` est incorporé dans une fonction de renvoi de promesse.</span><span class="sxs-lookup"><span data-stu-id="66ad8-146">The call to `getSelectedDataAsync` is embedded in a promise-returning function.</span></span> <span data-ttu-id="66ad8-147">Pour plus d'informations sur la raison et la façon de le faire, voir [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span><span class="sxs-lookup"><span data-stu-id="66ad8-147">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="66ad8-148">`getSelectedDataAsync` renvoie un tableau car plusieurs diapositives peuvent être sélectionnées.</span><span class="sxs-lookup"><span data-stu-id="66ad8-148">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="66ad8-149">Dans ce scénario, l'utilisateur n'en a sélectionné qu'une seule, de sorte que le code obtient la première (0e) diapositive, qui est la seule sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="66ad8-149">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="66ad8-150">La valeur de la diapositive est la valeur 1 que l'utilisateur voit en regard de la diapositive dans le volet miniatures de l'interface utilisateur `index` PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="66ad8-150">The `index` value of the slide is the 1-based value the user sees beside the slide in the PowerPoint UI thumbnails pane.</span></span>

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

3. <span data-ttu-id="66ad8-151">Le code suivant crée une méthode pour supprimer les diapositives marquées pour les clients premium.</span><span class="sxs-lookup"><span data-stu-id="66ad8-151">The following code creates a method to delete slides that are tagged for premium customers.</span></span> <span data-ttu-id="66ad8-152">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="66ad8-152">About this code, note:</span></span>

    - <span data-ttu-id="66ad8-153">Étant donné `key` que les `value` propriétés des balises vont être lues après `context.sync` le , elles doivent être chargées en premier.</span><span class="sxs-lookup"><span data-stu-id="66ad8-153">Because the `key` and `value` properties of the tags are going to be read after the `context.sync`, they must be loaded first.</span></span>

    ```javascript
    async function deleteSlidesByAudience() {
      await PowerPoint.run(async function(context) {
        const slides = context.presentation.slides;
        slides.load("tags/key, tags/value");
    
        await context.sync();
    
        for (let i = 0; i < slides.items.length; i++) {
          let currentSlide = slides.items[i];
          for (let j = 0; j < currentSlide.tags.items.length; j++) {
            let currentTag = currentSlide.tags.items[j];
            if (currentTag.key === "CUSTOMER_TYPE" && currentTag.value === "Premium") {
              currentSlide.delete();
            }
          }
        }
    
        await context.sync();
      });
    }
    ```

## <a name="set-custom-metadata-on-the-presentation"></a><span data-ttu-id="66ad8-154">Définir des métadonnées personnalisées sur la présentation</span><span class="sxs-lookup"><span data-stu-id="66ad8-154">Set custom metadata on the presentation</span></span>

<span data-ttu-id="66ad8-155">Les add-ins peuvent également appliquer des balises à la présentation dans son ensemble.</span><span class="sxs-lookup"><span data-stu-id="66ad8-155">Add-ins can also apply tags to the presentation as a whole.</span></span> <span data-ttu-id="66ad8-156">Cela vous permet d'utiliser des balises pour les métadonnées au niveau du document, de la même manière que la classe [CustomProperty](/javascript/api/word/word.customproperty)est utilisée dans Word.</span><span class="sxs-lookup"><span data-stu-id="66ad8-156">This enables you to use tags for document-level metadata similar to how the [CustomProperty](/javascript/api/word/word.customproperty)class is used in Word.</span></span> <span data-ttu-id="66ad8-157">Toutefois, contrairement à `CustomProperty` la classe Word, la valeur d'une balise PowerPoint peut uniquement être de type `string` .</span><span class="sxs-lookup"><span data-stu-id="66ad8-157">But unlike the Word `CustomProperty` class, the value of a PowerPoint tag can only be of type `string`.</span></span>

<span data-ttu-id="66ad8-158">Le code suivant est un exemple d'ajout d'une balise à une présentation.</span><span class="sxs-lookup"><span data-stu-id="66ad8-158">The following code is an example of adding a tag to a presentation.</span></span> 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

<span data-ttu-id="66ad8-159">Le code suivant est un exemple de suppression d'une balise d'une présentation.</span><span class="sxs-lookup"><span data-stu-id="66ad8-159">The following code is an example of deleting a tag from a presentation.</span></span> <span data-ttu-id="66ad8-160">Notez que la clé de la balise est transmise à la `delete` méthode de l'objet `TagsCollection` parent.</span><span class="sxs-lookup"><span data-stu-id="66ad8-160">Note that the key of the tag is passed to the `delete` method of the parent `TagsCollection` object.</span></span>

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
