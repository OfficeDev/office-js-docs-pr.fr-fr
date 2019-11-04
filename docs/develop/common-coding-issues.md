---
title: Problèmes de codage courants et comportements de plateforme inattendus
description: Liste des problèmes de plateforme d’API JavaScript pour Office fréquemment rencontrés par les développeurs.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d39c379961833cdb924628becf2c2da3f7e271b9
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924793"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a>Problèmes de codage courants et comportements de plateforme inattendus

Cet article met en évidence les aspects de l’API JavaScript pour Office qui peuvent entraîner un comportement inattendu ou nécessiter des modèles de codage spécifiques pour obtenir le résultat souhaité. Si vous rencontrez un problème qui se trouve dans cette liste, faites-le nous connaître en utilisant le formulaire de commentaires au bas de l’article.

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a>Les API communes et les API Outlook ne sont pas basées sur la promesse

Les [API communes](/javascript/api/office) (qui ne sont pas liées à un hôte Office particulier) et les [API Outlook](/javascript/api/outlook) utilisent un modèle de programmation basé sur les rappels. L’interaction avec le document Office sous-jacent nécessite un appel asynchrone en lecture ou en écriture qui spécifie un rappel à exécuter lorsque l’opération se termine. Pour obtenir un exemple de ce modèle, consultez la rubrique [document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).

Ces méthodes d’API et d’API courantes ne renvoient pas de [promesses](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise). Par conséquent, vous ne pouvez pas utiliser [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) pour suspendre l’exécution jusqu’à la fin de l’opération asynchrone. Si vous avez `await` besoin de comportement, vous pouvez encapsuler l’appel de méthode dans une promesse créée de manière explicite.

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> La documentation de référence contient l’implémentation encapsulée de [fichier. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).

## <a name="some-properties-must-be-set-with-json-structs"></a>Certaines propriétés doivent être définies avec des structs JSON

> [!NOTE]
> Cette section s’applique uniquement aux API propres à l’hôte pour Excel et Word.

Certaines propriétés doivent être définies en tant que structs JSON, au lieu de définir leurs sous-propriétés individuelles. Vous trouverez un exemple dans [PageLayout](/javascript/api/excel/excel.pagelayout). La `zoom` propriété doit être définie avec un seul objet [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , comme illustré ci-dessous :

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

Dans l’exemple précédent, vous ne seriez ***pas*** en mesure d' `zoom` affecter directement une `sheet.pageLayout.zoom.scale = 200;`valeur :. Cette instruction génère une erreur car `zoom` elle n’est pas chargée. Même si `zoom` elles ont été chargées, l’ensemble de l’étendue ne prendra pas effet. Toutes les opérations de `zoom`contexte se produisent, actualisant l’objet proxy dans le complément et remplaçant les valeurs définies localement.

Ce comportement diffère des [Propriétés de navigation](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) telles que [Range. format](/javascript/api/excel/excel.range#format). Les propriétés `format` de peuvent être définies à l’aide de la navigation d’objet, comme illustré ci-dessous :

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Vous pouvez identifier une propriété dont les propriétés subordonnées doivent être définies avec un struct JSON en vérifiant son modificateur en lecture seule. Les propriétés non en lecture seule de toutes les propriétés en lecture seule peuvent être définies directement. Les propriétés accessibles en `PageLayout.zoom` écriture comme doivent être définies avec une structure JSON. En Résumé :

- Propriété en lecture seule : les sous-propriétés peuvent être définies via la navigation.
- Propriété accessible en écriture : les sous-propriétés doivent être définies avec une structure JSON (et ne peuvent pas être définies via la navigation).

## <a name="excel-range-limits"></a>Limites de plage Excel

Si vous créez un complément Excel qui utilise des plages, gardez à l’esprit les limitations de taille suivantes :

- Excel sur le web a une limite de taille de charge utile de 5 Mo pour les demandes et les réponses. L’erreur `RichAPI.Error` est déclenchée en cas de dépassement de cette limite.
- Une plage est limitée à 5 millions cellules pour les opérations Set.

Si vous prévoyez que l’entrée de l’utilisateur dépasse ces limites, veillez à vérifier les données et à les fractionner en plusieurs objets. Vous devrez également envoyer plusieurs `context.sync()` appels afin d’éviter que les opérations de plage plus petites soient regroupées.

Votre complément peut utiliser [RangeAreas](/javascript/api/excel/excel.rangeareas) pour mettre à jour les cellules dans une plage plus grande de manière stratégique. Pour plus d’informations, consultez [travailler simultanément avec plusieurs plages dans des compléments Excel](../excel/excel-add-ins-multiple-ranges.md) .

## <a name="setting-read-only-properties"></a>Définition de propriétés en lecture seule

Les [définitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) de la machine à écrire pour Office js spécifient les propriétés d’objet en lecture seule. Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue sans avertissement, sans qu’aucune erreur ne soit générée. L’exemple suivant tente à tort de définir la propriété en lecture seule [Chart.ID](/javascript/api/excel/excel.chart#id).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a>Voir aussi

- [OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): le lieu de signaler et d’afficher les problèmes liés à la plateforme des compléments Office et aux API JavaScript.
- [Débordement de pile](https://stackoverflow.com/questions/tagged/office-js): emplacement où poser des questions de programmation sur les API JavaScript Office. Veillez à appliquer la balise « Office-js » à votre question lors de la publication dans le débordement de pile.
- [UserVoice](https://officespdev.uservoice.com/): le lieu de suggérer de nouvelles fonctionnalités pour la plateforme des compléments Office et les API JavaScript pour Office.
