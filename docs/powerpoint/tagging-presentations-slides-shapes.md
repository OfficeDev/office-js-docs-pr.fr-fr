---
title: Utiliser des balises personnalisées sur des présentations, des diapositives et des formes dans PowerPoint
description: Découvrez comment utiliser des balises pour les métadonnées personnalisées sur les présentations, les diapositives et les formes.
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: a30beea56286437b1c69461534ca13912107cecf
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958901"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a>Utiliser des balises personnalisées pour les présentations, les diapositives et les formes dans PowerPoint

Un complément peut attacher des métadonnées personnalisées, sous la forme de paires clé-valeur, appelées « balises », à des présentations, à des diapositives spécifiques et à des formes spécifiques sur une diapositive.

Il existe deux scénarios principaux pour l’utilisation de balises :

- Lorsqu’elle est appliquée à une diapositive ou à une forme, une balise permet de catégoriser l’objet pour le traitement par lots. Par exemple, supposons qu’une présentation comporte des diapositives qui doivent être incluses dans les présentations dans la région Est, mais pas dans la région Ouest. De même, il existe d’autres diapositives qui doivent être affichées uniquement à l’Ouest. Votre complément peut créer une balise avec la clé `REGION` et la valeur `East` et l’appliquer aux diapositives qui ne doivent être utilisées qu’à l’est. La valeur de la balise est définie `West` pour les diapositives qui ne doivent être affichées que dans la région Ouest. Juste avant une présentation à l’Est, un bouton du complément exécute du code qui effectue une boucle dans toutes les diapositives en vérifiant la valeur de la `REGION` balise. Diapositives dans lesquelles la région est `West` supprimée. L’utilisateur ferme ensuite le complément et démarre le diaporama.
- Lorsqu’elle est appliquée à une présentation, une balise est en fait une propriété personnalisée dans le document de présentation (similaire à une [customProperty](/javascript/api/word/word.customproperty) dans Word).

## <a name="tag-slides-and-shapes"></a>Baliser des diapositives et des formes

Une balise est une paire clé-valeur, où la valeur est toujours de type `string` et est représentée par un objet [Tag](/javascript/api/powerpoint/powerpoint.tag) . Chaque type d’objet parent, tel qu’un objet [Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide) ou [Shape](/javascript/api/powerpoint/powerpoint.shape) , a une `tags` propriété de type [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).

### <a name="add-update-and-delete-tags"></a>Ajouter, mettre à jour et supprimer des balises

Pour ajouter une balise à un objet, appelez la méthode [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1)) de la propriété de `tags` l’objet parent. Le code suivant ajoute deux balises à la première diapositive d’une présentation. Tenez compte du code suivant :

- Le premier paramètre de la `add` méthode est la clé dans la paire clé-valeur.
- Le deuxième paramètre est la valeur.
- La clé est en lettres majuscules. Cela n’est pas strictement obligatoire pour la `add` méthode. Toutefois, la clé est toujours stockée par PowerPoint en majuscules, et *certaines méthodes liées aux balises nécessitent que la clé soit exprimée en majuscules*. Nous vous recommandons donc de toujours utiliser des majuscules dans votre code pour une clé de balise.

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

Pour supprimer une balise, appelez la `delete` méthode sur son objet parent `TagsCollection` et passez la clé de la balise en tant que paramètre. Pour obtenir un exemple, consultez [Définir des métadonnées personnalisées sur la présentation](#set-custom-metadata-on-the-presentation).

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a>Utiliser des balises pour traiter de manière sélective les diapositives et les formes

Prenons le scénario suivant : Contoso Consulting propose une présentation qu’il présente à tous les nouveaux clients. Toutefois, certaines diapositives ne doivent être affichées qu’aux clients qui ont payé pour l’état « Premium ». Avant d’afficher la présentation aux clients non Premium, ils en font une copie et suppriment les diapositives que seuls les clients Premium doivent voir. Un complément permet à Contoso de baliser les diapositives destinées aux clients Premium et de supprimer ces diapositives si nécessaire. La liste suivante décrit les principales étapes de codage pour créer cette fonctionnalité.

1. Créez une fonction qui balise la diapositive actuellement sélectionnée comme prévu pour les `Premium` clients. Tenez compte du code suivant :

    - La `getSelectedSlideIndex` fonction est définie à l’étape suivante. Elle retourne l’index basé sur 1 de la diapositive actuellement sélectionnée.
    - La valeur retournée par la `getSelectedSlideIndex` fonction doit être décrémentée, car la méthode [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1)) est basée sur 0.

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

    - Il utilise la méthode [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) des API JavaScript courantes.
    - L’appel à `getSelectedDataAsync` est incorporé dans une fonction de retour de promesse. Pour plus d’informations sur la raison et la procédure à suivre, consultez [Wrap Common API in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).
    - `getSelectedDataAsync` retourne un tableau, car plusieurs diapositives peuvent être sélectionnées. Dans ce scénario, l’utilisateur n’a sélectionné qu’une seule diapositive, de sorte que le code obtient la première (0e) diapositive, qui est la seule sélectionnée.
    - La `index` valeur de la diapositive est la valeur basée sur 1 que l’utilisateur voit à côté de la diapositive dans le volet miniatures de l’interface utilisateur PowerPoint.

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

3. Le code suivant crée une fonction pour supprimer les diapositives marquées pour les clients Premium. Tenez compte du code suivant :

    - Étant donné que les propriétés et `value` les `key` balises vont être lues après la `context.sync`balise, elles doivent d’abord être chargées.

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

Les compléments peuvent également appliquer des balises à la présentation dans son ensemble. Cela vous permet d’utiliser des balises pour les métadonnées au niveau du document, comme la classe [CustomProperty](/javascript/api/word/word.customproperty)est utilisée dans Word. Mais contrairement à la classe Word `CustomProperty` , la valeur d’une balise PowerPoint ne peut être que de type `string`.

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

Le code suivant est un exemple de suppression d’une balise d’une présentation. Notez que la clé de la balise est passée à la `delete` méthode de l’objet parent `TagsCollection` .

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
