---
title: Ajouter et supprimer des diapositives dans PowerPoint
description: Découvrez comment ajouter et supprimer des diapositives et spécifier le maître et la disposition des nouvelles diapositives.
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2cf22c18cf4089bab9091be3f4274f67974662a3
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958312"
---
# <a name="add-and-delete-slides-in-powerpoint"></a>Ajouter et supprimer des diapositives dans PowerPoint

Un complément PowerPoint peut ajouter des diapositives à la présentation et spécifier éventuellement le masque des diapositives et la disposition du masque utilisés pour la nouvelle diapositive. Le complément peut également supprimer des diapositives.

Les API permettant d’ajouter des diapositives sont principalement utilisées dans les scénarios où les ID des masques des diapositives et des dispositions de la présentation sont connus au moment du codage ou se trouvent dans une source de données au moment de l’exécution. Dans un tel scénario, vous ou le client devez créer et gérer une source de données qui met en corrélation le critère de sélection (par exemple, les noms ou les images des masques des diapositives et des dispositions) avec les ID des masques de diapositives et des dispositions. Les API peuvent également être utilisées dans les scénarios où l’utilisateur peut insérer des diapositives qui utilisent le masque des diapositives par défaut et la disposition par défaut du masque, et dans les scénarios où l’utilisateur peut sélectionner une diapositive existante et en créer une avec le même masque des diapositives et la même disposition (mais pas le même contenu). Pour plus [d’informations à ce sujet, consultez Sélection du masque des diapositives et de la disposition à utiliser](#select-which-slide-master-and-layout-to-use) .

## <a name="add-a-slide-with-slidecollectionadd"></a>Ajouter une diapositive avec SlideCollection.add

Ajoutez des diapositives avec la méthode [SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1)) . Voici un exemple simple dans lequel une diapositive qui utilise le masque des diapositives par défaut de la présentation et la première disposition de ce masque sont ajoutées. La méthode ajoute toujours de nouvelles diapositives à la fin de la présentation. Voici un exemple.

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="select-which-slide-master-and-layout-to-use"></a>Sélectionner le masque des diapositives et la disposition à utiliser

Utilisez le paramètre [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) pour contrôler quel masque des diapositives est utilisé pour la nouvelle diapositive et quelle disposition dans le masque est utilisée. Voici un exemple. Tenez compte du code suivant :

- Vous pouvez inclure l’une ou les deux propriétés de l’objet `AddSlideOptions` .
- Si les deux propriétés sont utilisées, la disposition spécifiée doit appartenir au maître spécifié ou une erreur est levée.
- Si la `masterId` propriété n’est pas présente (ou si sa valeur est une chaîne vide), le masque des diapositives par défaut est utilisé et doit `layoutId` être une disposition de ce masque des diapositives.
- Le masque des diapositives par défaut est le masque des diapositives utilisé par la dernière diapositive de la présentation. (Dans le cas inhabituel où il n’y a actuellement aucune diapositive dans la présentation, le masque des diapositives par défaut est le premier masque des diapositives de la présentation.)
- Si la `layoutId` propriété n’est pas présente (ou si sa valeur est une chaîne vide), la première disposition du maître spécifié par l’objet `masterId` est utilisée.
- Les deux propriétés sont des chaînes d’une des trois formes possibles : ***nnnnnnnnnn*#**, **#* mmmmmmmmm**, ou **_nnnnnnnnnn_#* mmmmmmm***, où *nnnnnnnnnn* est l’ID du maître ou de la disposition (généralement 10 chiffres) et *mmmmmmmmm est* l’ID de création du maître ou de la disposition (généralement 6 à 10 chiffres). Voici quelques exemples : `2147483690#2908289500`, `2147483690#`et `#2908289500`.

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

Il n’existe aucun moyen pratique pour les utilisateurs de découvrir l’ID ou l’ID de création d’un masque des diapositives ou d’une disposition. Pour cette raison, vous pouvez vraiment utiliser le `AddSlideOptions` paramètre uniquement lorsque vous connaissez les ID au moment du codage ou que votre complément peut les découvrir au moment de l’exécution. Étant donné que les utilisateurs ne peuvent pas être censés mémoriser les ID, vous avez également besoin d’un moyen de permettre à l’utilisateur de sélectionner des diapositives, par exemple par nom ou par une image, puis de mettre en corrélation chaque titre ou image avec l’ID de la diapositive.

En conséquence, le `AddSlideOptions` paramètre est principalement utilisé dans les scénarios dans lesquels le complément est conçu pour fonctionner avec un ensemble spécifique de masques de diapositives et de dispositions dont les ID sont connus. Dans ce scénario, vous ou le client devez créer et gérer une source de données qui met en corrélation un critère de sélection (par exemple, des noms ou des images de masque des diapositives et de disposition) avec les ID ou ID de création correspondants.

#### <a name="have-the-user-choose-a-matching-slide"></a>Faire choisir à l’utilisateur une diapositive correspondante

Si votre complément peut être utilisé dans les scénarios où la nouvelle diapositive doit utiliser la même combinaison de masque des diapositives et de disposition que celle utilisée par une diapositive *existante* , votre complément peut (1) inviter l’utilisateur à sélectionner une diapositive et (2) lire les ID du masque des diapositives et de la disposition. Les étapes suivantes montrent comment lire les ID et ajouter une diapositive avec une forme de base et une disposition correspondantes.

1. Créez une fonction pour obtenir l’index de la diapositive sélectionnée. Voici un exemple. Tenez compte du code suivant :

    - Il utilise la méthode [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) des API JavaScript courantes.
    - L’appel à `getSelectedDataAsync` est incorporé dans une fonction de retour de promesse. Pour plus d’informations sur la raison et la procédure à suivre, consultez [Wrap Common API in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).
    - `getSelectedDataAsync` retourne un tableau, car plusieurs diapositives peuvent être sélectionnées. Dans ce scénario, l’utilisateur n’a sélectionné qu’une seule diapositive, de sorte que le code obtient la première (0e) diapositive, qui est la seule sélectionnée.
    - La `index` valeur de la diapositive est la valeur basée sur 1 que l’utilisateur voit à côté de la diapositive dans le volet miniatures.

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

2. Appelez votre nouvelle fonction à l’intérieur de [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) de la fonction principale qui ajoute la diapositive. Voici un exemple.

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

## <a name="delete-slides"></a>Supprimer des diapositives

Supprimez une diapositive en obtenant une référence à l’objet [Slide](/javascript/api/powerpoint/powerpoint.slide) qui représente la diapositive et appelez la `Slide.delete` méthode. Voici un exemple dans lequel la 4e diapositive est supprimée.

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
