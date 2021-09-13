---
title: Ajouter et supprimer des diapositives dans PowerPoint
description: Découvrez comment ajouter et supprimer des diapositives et spécifier le maître et la mise en page des nouvelles diapositives.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 26999ed770fa8fde8766a2accb7ec9eb791fb3d4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150404"
---
# <a name="add-and-delete-slides-in-powerpoint"></a>Ajouter et supprimer des diapositives dans PowerPoint

Un PowerPoint peut ajouter des diapositives à la présentation et éventuellement spécifier le maître des diapositives et la mise en page du maître utilisé pour la nouvelle diapositive. Le add-in peut également supprimer des diapositives.

> [!IMPORTANT]
> Les API d’ajout de diapositives sont en [prévisualisation](../reference/requirement-sets/powerpoint-preview-apis.md) et ne sont pas disponibles pour les modules de production. L’API *de suppression des* diapositives a été publiée.

Les API d’ajout de diapositives sont principalement utilisées dans les scénarios où les ID des formes de base et des mises en page des diapositives de la présentation sont connus au moment du codage ou se trouvent dans une source de données lors de l’runtime. Dans ce cas, vous ou le client devez créer et gérer une source de données qui met en corrélation le critère de sélection (par exemple, les noms ou les images des formes de base et des mises en page des diapositives) avec les ID des formes de base et des mises en page des diapositives. Les API peuvent également être utilisées dans les scénarios où l’utilisateur peut insérer des diapositives qui utilisent le maître des diapositives par défaut et la mise en page par défaut du maître, et dans les scénarios où l’utilisateur peut sélectionner une diapositive existante et en créer une nouvelle avec le même maître et la même mise en page de diapositives (mais pas le même contenu). Pour [plus d’informations à](#select-which-slide-master-and-layout-to-use) ce sujet, voir Sélection du maître des diapositives et de la mise en page à utiliser.

## <a name="add-a-slide-with-slidecollectionadd-preview"></a>Ajouter une diapositive avec SlideCollection.add (aperçu)

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

Ajoutez des diapositives avec [la méthode SlideCollection.add.](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) Voici un exemple simple dans lequel une diapositive qui utilise le maître des diapositives par défaut de la présentation et la première mise en page de ce maître est ajoutée. La méthode ajoute toujours de nouvelles diapositives à la fin de la présentation. Voici un exemple.

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="select-which-slide-master-and-layout-to-use"></a>Sélectionnez le maître des diapositives et la mise en page à utiliser

Utilisez le [paramètre AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) pour contrôler le maître des diapositives qui est utilisé pour la nouvelle diapositive et la mise en page dans le master. Voici un exemple. Tenez compte du code suivant :

- Vous pouvez inclure l’une ou l’autre des propriétés de l’objet ou les `AddSlideOptions` deux.
- Si les deux propriétés sont utilisées, la disposition spécifiée doit appartenir à la forme de base spécifiée ou une erreur est lancée.
- Si la propriété n’est pas présente (ou si sa valeur est une chaîne vide), le curseur de diapositive par défaut est utilisé et doit être une mise en page de `masterId` `layoutId` ce dernier.
- Le maître des diapositives par défaut est celui utilisé par la dernière diapositive de la présentation. (Dans le cas rare où il n’y a actuellement aucune diapositive dans la présentation, le maître des diapositives par défaut est le premier maître des diapositives de la présentation.)
- Si la propriété n’est pas présente (ou si sa valeur est une chaîne vide), la première disposition de la forme de base spécifiée par la forme de base `layoutId` `masterId` est utilisée.
- Les deux propriétés sont des chaînes de l’une des trois formes possibles : ***nnnnnnnnnn*#**, * *#* mmmmmmmmmmm*** ou **_nnnnnnnnnn_ #* mmmmmmmmm***, où *nnnnnnnnnn* est l’ID de la forme de base ou de la disposition (généralement 10 chiffres) et *mmmmmmmmmmm* est l’ID de création de la forme de base ou de la disposition (généralement 6 à 10 chiffres). Voici quelques exemples `2147483690#2908289500` : `2147483690#` , et `#2908289500` .

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

Il n’existe aucun moyen pratique pour les utilisateurs de découvrir l’ID ou l’ID de création d’un curseur de diapositive ou d’une mise en page. Pour cette raison, vous ne pouvez utiliser le paramètre que lorsque vous connaissez les ID au moment du codage ou que votre application peut les découvrir lors de `AddSlideOptions` l’utilisation. Étant donné que les utilisateurs ne sont pas censés mémoriser les ID, vous avez également besoin d’un moyen pour permettre à l’utilisateur de sélectionner des diapositives, par exemple par son nom ou par une image, puis de corréler chaque titre ou image avec l’ID de la diapositive.

Par conséquent, le paramètre est principalement utilisé dans les scénarios dans lesquels le module est conçu pour fonctionner avec un ensemble spécifique de formes de base et de mises en page dont les ID sont `AddSlideOptions` connus. Dans ce cas, vous ou le client devez créer et gérer une source de données qui met en corrélation un critère de sélection (tel que le maître des diapositives et les noms ou images de mise en page) avec les ID ou les ID de création correspondants.

#### <a name="have-the-user-choose-a-matching-slide"></a>Faire en cas de choix d’une diapositive correspondante par l’utilisateur

Si votre add-in peut être utilisé dans des scénarios où la nouvelle diapositive doit  utiliser la même combinaison de formes de base et de mise en page que celle utilisée par une diapositive existante, votre add-in peut (1) invite l’utilisateur à sélectionner une diapositive et (2) lit les ID du maître et de la mise en page des diapositives. Les étapes suivantes montrent comment lire les ID et ajouter une diapositive avec une forme de base et une mise en page correspondantes.

1. Créez une méthode pour obtenir l’index de la diapositive sélectionnée. Voici un exemple. Tenez compte du code suivant :

    - Il utilise la [méthode Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) des API JavaScript communes.
    - L’appel `getSelectedDataAsync` est incorporé dans une fonction de renvoi de promesse. Pour plus d’informations sur la raison et la façon de le faire, voir [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).
    - `getSelectedDataAsync` renvoie un tableau car plusieurs diapositives peuvent être sélectionnées. Dans ce scénario, l’utilisateur n’en a sélectionné qu’une seule, de sorte que le code obtient la première (0e) diapositive, qui est la seule sélectionnée.
    - La valeur de la diapositive est la valeur 1 que l’utilisateur voit en regard de la diapositive dans le volet de `index` miniatures.

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

2. Appelez votre nouvelle fonction à [l’intérieur PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) de la fonction principale qui ajoute la diapositive. Voici un exemple.

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

Supprimez une diapositive en obtenant une référence à l’objet [Slide](/javascript/api/powerpoint/powerpoint.slide) qui représente la diapositive et appelez la `Slide.delete` méthode. Voici un exemple dans lequel la quatrième diapositive est supprimée.

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
