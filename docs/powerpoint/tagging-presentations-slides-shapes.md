---
title: 'Utiliser des balises personnalisées sur les présentations, diapositives et formes dans PowerPoint'
description: 'Découvrez comment utiliser des balises pour des métadonnées personnalisées sur les présentations, les diapositives et les formes.'
ms.date: 12/14/2021
ms.localizationpriority: medium
---

# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a>Utiliser des balises personnalisées pour les présentations, diapositives et formes dans PowerPoint

Un add-in peut joindre des métadonnées personnalisées, sous la forme de paires clé-valeur, appelées « balises », à des présentations, des diapositives spécifiques et des formes spécifiques sur une diapositive.

Il existe deux scénarios principaux pour l’utilisation de balises :

- Lorsqu’elle est appliquée à une diapositive ou à une forme, une balise permet de classer l’objet pour le traitement par lots. Par exemple, supposons qu’une présentation possède des diapositives qui doivent être incluses dans les présentations de la région Est, mais pas de la région Ouest. De même, il existe d’autres diapositives qui doivent être affichées uniquement à l’Ouest. Votre application peut créer `REGION` `East` une balise avec la clé et la valeur et l’appliquer aux diapositives qui ne doivent être utilisées qu’à l’Est. La valeur de la balise est définie pour `West` les diapositives qui doivent uniquement être affichées dans la région Ouest. Juste avant une présentation à l’Est, un bouton du add-in exécute un code qui pare toutes les diapositives en vérifiant la valeur de la `REGION` balise. Diapositives dans laquelle la région est `West` supprimée. L’utilisateur ferme ensuite le module et démarre le diaporama.
- Lorsqu’elle est appliquée à une présentation, une balise est en fait une propriété personnalisée dans le document de présentation (semblable à [une propriété](/javascript/api/word/word.customproperty) personnalisée dans Word).

## <a name="tag-slides-and-shapes"></a>Baliser les diapositives et les formes

Une balise est une paire clé-valeur, où la valeur est toujours de type `string` et est représentée par un [objet Tag](/javascript/api/powerpoint/powerpoint.tag) . Chaque type d’objet parent, tel qu’un objet [Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide) ou [Shape](/javascript/api/powerpoint/powerpoint.shape) , possède une `tags` propriété de type [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).

### <a name="add-update-and-delete-tags"></a>Ajouter, mettre à jour et supprimer des balises

Pour ajouter une balise à un objet, appelez la méthode [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1)) de la propriété de l’objet `tags` parent. Le code suivant ajoute deux balises à la première diapositive d’une présentation. Tenez compte du code suivant :

- Le premier paramètre de la méthode `add` est la clé de la paire clé-valeur.
- Le deuxième paramètre est la valeur.
- La clé est en lettres majuscules. Cela n’est `add` pas strictement obligatoire pour la méthode ; toutefois, la clé est toujours stockée par PowerPoint en tant que minuscules, et certaines méthodes *liées aux balises* nécessitent que la clé soit exprimée en minuscules. Nous vous recommandons donc de toujours utiliser des minuscules dans votre code pour une clé de balise.

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

La `add` méthode est également utilisée pour mettre à jour une balise. Le code suivant modifie la valeur de la `PLANET` balise.

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

Pour supprimer une balise, appelez la `delete` méthode sur son objet parent `TagsCollection` et passez la clé de la balise en tant que paramètre. Pour obtenir un exemple, voir [Définir des métadonnées personnalisées dans la présentation](#set-custom-metadata-on-the-presentation).

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a>Utiliser des balises pour traiter de manière sélective les diapositives et les formes

Envisagez le scénario suivant : Contoso Consulting présente une présentation qu’il présente à tous les nouveaux clients. Toutefois, certaines diapositives ne doivent être affichées qu’aux clients qui ont payé l’état « premium ». Avant d’afficher la présentation aux clients non premium, ils en font une copie et suppriment les diapositives que seuls les clients premium doivent voir. Un add-in permet à Contoso de baliser les diapositives qui sont pour les clients premium et de supprimer ces diapositives si nécessaire. La liste suivante décrit les principales étapes de codage pour créer cette fonctionnalité.

1. Créez une méthode qui balise la diapositive actuellement sélectionnée comme prévu pour les `Premium` clients. Tenez compte du code suivant :

    - La `getSelectedSlideIndex` fonction est définie à l’étape suivante. Elle renvoie l’index de base 1 de la diapositive actuellement sélectionnée.
    - La valeur renvoyée par la `getSelectedSlideIndex` fonction doit être décrémentée car la méthode [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1)) est basée sur 0.

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

2. Le code suivant crée une méthode pour obtenir l’index de la diapositive sélectionnée. Tenez compte du code suivant :

    - Il utilise la [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) des API JavaScript communes.
    - L’appel est `getSelectedDataAsync` incorporé dans une fonction de renvoi de promesse. Pour plus d’informations sur la raison et la façon de le faire, voir [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).
    - `getSelectedDataAsync` renvoie un tableau car plusieurs diapositives peuvent être sélectionnées. Dans ce scénario, l’utilisateur n’en a sélectionné qu’une seule, de sorte que le code obtient la première (0e) diapositive, qui est la seule sélectionnée.
    - La `index` valeur de la diapositive est la valeur basée sur 1 que l’utilisateur voit en regard de la diapositive dans le PowerPoint miniatures de l’interface utilisateur.

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

3. Le code suivant crée une méthode pour supprimer les diapositives marquées pour les clients premium. Tenez compte du code suivant :

    - Étant donné que `key` les propriétés `value` des balises vont être lues après `context.sync`le , elles doivent être chargées en premier.

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

## <a name="set-custom-metadata-on-the-presentation"></a>Définir des métadonnées personnalisées sur la présentation

Les add-ins peuvent également appliquer des balises à la présentation dans son ensemble. Cela vous permet d’utiliser des balises pour les métadonnées au niveau du document, de la même façon que la [classe CustomProperty](/javascript/api/word/word.customproperty) est utilisée dans Word. Toutefois, contrairement à la classe Word`CustomProperty`, la valeur d’une balise PowerPoint ne peut être que de type `string`.

Le code suivant est un exemple d’ajout d’une balise à une présentation. 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

Le code suivant est un exemple de suppression d’une balise d’une présentation. Notez que la clé de la balise est transmise à la `delete` méthode de l’objet parent `TagsCollection` .

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
