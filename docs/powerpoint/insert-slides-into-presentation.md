---
title: Insérer des diapositives dans une présentation PowerPoint présentation
description: Découvrez comment insérer des diapositives d’une présentation dans une autre.
ms.date: 03/07/2021
ms.localizationpriority: medium
ms.openlocfilehash: c7dde2d2d6b1b886816bbf12122319984f4c7138
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152202"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a>Insérer des diapositives dans une présentation PowerPoint présentation

Un PowerPoint peut insérer des diapositives d’une présentation dans la présentation actuelle à l’aide PowerPoint bibliothèque JavaScript propre à l’application. Vous pouvez contrôler si les diapositives insérées conservent la mise en forme de la présentation source ou la mise en forme de la présentation cible.

Les API d’insertion de diapositives sont principalement utilisées dans les scénarios de modèles de présentation : il existe un petit nombre de présentations connues qui servent de pools de diapositives qui peuvent être insérées par le module. Dans ce cas, vous ou le client devez créer et gérer une source de données qui met en corrélation le critère de sélection (par exemple, titres ou images) avec les ID de diapositive. Les API peuvent également être utilisées dans des scénarios où l’utilisateur peut insérer des diapositives  à partir de n’importe quelle présentation arbitraire, mais dans ce scénario, l’utilisateur est effectivement limité à l’insertion de toutes les diapositives de la présentation source. Pour [plus d’informations à](#selecting-which-slides-to-insert) ce sujet, voir Sélection des diapositives à insérer.

Il existe deux étapes pour insérer des diapositives d’une présentation dans une autre.

1. Convertissez le fichier de présentation source (.pptx) en chaîne au format Base64.
1. Utilisez la méthode pour insérer une ou plusieurs diapositives du `insertSlidesFromBase64` fichier Base64 dans la présentation actuelle.

## <a name="convert-the-source-presentation-to-base64"></a>Convertir la présentation source en base64

Il existe plusieurs façons de convertir un fichier en base64. Le langage de programmation et la bibliothèque que vous utilisez, et s’il faut les convertir côté serveur ou côté client, sont déterminés par votre scénario. Le plus souvent, vous allez faire la conversion dans JavaScript côté client à l’aide d’un [objet FileReader.](https://developer.mozilla.org/docs/Web/API/FileReader) L’exemple suivant illustre cette pratique.

1. Commencez par obtenir une référence au fichier PowerPoint source. Dans cet exemple, nous allons utiliser un contrôle de type pour demander à `<input>` l’utilisateur de choisir un `file` fichier. Ajoutez le markup suivant à la page du add-in.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    Ce markup ajoute l’interface utilisateur dans la capture d’écran suivante à la page.

    ![Screenshot showing an HTML file type input control preceded by an instructional sentence reading « Select a PowerPoint presentation from which to insert slides ». Le contrôle se compose d’un bouton étiqueté « Choisir un fichier » suivi de la phrase « Aucun fichier choisi ».](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > Il existe de nombreuses autres façons d’obtenir un PowerPoint de données. Par exemple, si le fichier est stocké sur OneDrive ou SharePoint, vous pouvez utiliser Microsoft Graph pour le télécharger. Pour plus d’informations, voir [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and Access Files with Microsoft [Graph](/learn/modules/msgraph-access-file-data/).

2. Ajoutez le code suivant au code JavaScript du add-in pour affecter une fonction à l’événement du contrôle `change` d’entrée. (Vous créez la `storeFileAsBase64` fonction à l’étape suivante.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Ajoutez le code suivant. Notez ce qui suit à propos de ce code.

    - La `reader.readAsDataURL` méthode convertit le fichier en base64 et le stocke dans la `reader.result` propriété. Une fois la méthode terminée, elle déclenche le `onload` handler d’événements.
    - Le handler d’événements coupe les métadonnées du fichier codé et stocke la chaîne codée `onload` dans une variable globale.
    - La chaîne codée en base 64 est stockée globalement, car elle sera lue par une autre fonction que vous créerez à une étape ultérieure.

    ```javascript
    let chosenFileBase64;

    async function storeFileAsBase64() {
        const reader = new FileReader();

        reader.onload = async (event) => {
            const startIndex = reader.result.toString().indexOf("base64,");
            const copyBase64 = reader.result.toString().substr(startIndex + 7);

            chosenFileBase64 = copyBase64;
        };

        const myFile = document.getElementById("file") as HTMLInputElement;
        reader.readAsDataURL(myFile.files[0]);
    }
    ```

## <a name="insert-slides-with-insertslidesfrombase64"></a>Insérer des diapositives avec insertSlidesFromBase64

Votre add-in insère des diapositives d’une autre PowerPoint présentation dans la présentation actuelle à l’aide de la méthode [Presentation.insertSlidesFromBase64.](/javascript/api/powerpoint/powerpoint.presentation#insertSlidesFromBase64_base64File__options_) Voici un exemple simple dans lequel toutes les diapositives de la présentation source sont insérées au début de la présentation en cours et les diapositives insérées conservent la mise en forme du fichier source. Notez qu’il s’agit d’une variable globale qui contient une version codée `chosenFileBase64` en base 64 d’PowerPoint de présentation.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

Vous pouvez contrôler certains aspects du résultat d’insertion, y compris l’endroit où les diapositives sont insérées et si elles obtiennent la mise en forme source ou cible, en passant un objet [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) en tant que deuxième paramètre à `insertSlidesFromBase64` . Voici un exemple. Tenez compte du code suivant :

- Il existe deux valeurs possibles pour la propriété `formatting` : « UseDestinationTheme » et « KeepSourceFormatting ». Si vous le souhaitez, vous pouvez utiliser `InsertSlideFormatting` l’enum (par exemple, `PowerPoint.InsertSlideFormatting.useDestinationTheme` ).
- La fonction insère les diapositives de la présentation source immédiatement après la diapositive spécifiée par la `targetSlideId` propriété. La valeur de cette propriété est une chaîne de l’une des trois formes possibles : ***nnn*#**, * *#* mmmmmmmmmmm*** ou **_nnn_ #* mmmmmmmmm***, où *nnn* est l’ID de la diapositive (généralement 3 chiffres) et *mmmmmmmmm est* l’ID de création de la diapositive (généralement 9 chiffres). Voici quelques exemples `267#763315295` : `267#` , et `#763315295` .

```javascript
async function insertSlidesDestinationFormatting() {
  await PowerPoint.run(async function(context) {
    context.presentation
    .insertSlidesFromBase64(chosenFileBase64,
                            {
                                formatting: "UseDestinationTheme",
                                targetSlideId: "267#"
                            }
                          );
    await context.sync();
  });
}
```

Bien entendu, vous ne connaissez généralement pas au moment du codage l’ID ou l’ID de création de la diapositive cible. Plus souvent, un add-in demande aux utilisateurs de sélectionner la diapositive cible. Les étapes suivantes montrent comment obtenir l’ID ***nnn*#** de la diapositive actuellement sélectionnée et l’utiliser comme diapositive cible.

1. Créez une fonction qui obtient l’ID de la diapositive actuellement sélectionnée à l’aide de la méthode [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) des API JavaScript communes. Voici un exemple. Notez que l’appel `getSelectedDataAsync` est incorporé dans une fonction de renvoi de promesse. Pour plus d’informations sur la raison et la façon de le faire, voir Wrap Common-APIs dans les fonctions [de renvoi de promesse.](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)

 
    ```javascript
    function getSelectedSlideID() {
      return new OfficeExtension.Promise<string>(function (resolve, reject) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
          try {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              reject(console.error(asyncResult.error.message));
            } else {
              resolve(asyncResult.value.slides[0].id);
            }
          }
          catch (error) {
            reject(console.log(error));
          }
        });
      })
    }
    ```

1. Appelez votre nouvelle fonction à l’intérieur de [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) de la fonction principale et passez l’ID qu’elle renvoie (concatentée avec le symbole « # » ) comme valeur de la propriété du `targetSlideId` `InsertSlideOptions` paramètre. Voici un exemple.

    ```javascript
    async function insertAfterSelectedSlide() {
        await PowerPoint.run(async function(context) {

            const selectedSlideID = await getSelectedSlideID();

            context.presentation.insertSlidesFromBase64(chosenFileBase64, {
                formatting: "UseDestinationTheme",
                targetSlideId: selectedSlideID + "#"
            });

            await context.sync();
        });
    }
    ```

### <a name="selecting-which-slides-to-insert"></a>Sélection des diapositives à insérer

Vous pouvez également utiliser le [paramètre InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) pour contrôler les diapositives de la présentation source qui sont insérées. Pour ce faire, affectez un tableau des ID de diapositive de la présentation source à la `sourceSlideIds` propriété. Voici un exemple qui insère quatre diapositives. Notez que chaque chaîne du tableau doit suivre l’un ou l’autre des modèles utilisés pour la `targetSlideId` propriété.

```javascript
async function insertAfterSelectedSlide() {
    await PowerPoint.run(async function(context) {
        const selectedSlideID = await getSelectedSlideID();
        context.presentation.insertSlidesFromBase64(chosenFileBase64, {
            formatting: "UseDestinationTheme",
            targetSlideId: selectedSlideID + "#",
            sourceSlideIds: ["267#763315295", "256#", "#926310875", "1270#"]
        });

        await context.sync();
    });
}
```

> [!NOTE]
> Les diapositives sont insérées dans le même ordre relatif dans lequel elles apparaissent dans la présentation source, quel que soit l’ordre dans lequel elles apparaissent dans le tableau.

Il n’existe aucun moyen pratique pour les utilisateurs de découvrir l’ID ou l’ID de création d’une diapositive dans la présentation source. Pour cette raison, vous ne pouvez utiliser la propriété que si vous connaissez les ID source au moment du codage ou que votre application peut les récupérer lors de l’utilisation à partir d’une source de `sourceSlideIds` données. Étant donné que les utilisateurs ne sont pas censés mémoriser les ID de diapositive, vous avez également besoin d’un moyen pour permettre à l’utilisateur de sélectionner des diapositives, par exemple par titre ou par une image, puis de corréler chaque titre ou image avec l’ID de la diapositive.

Par conséquent, la propriété est principalement utilisée dans les scénarios de modèles de présentation : le add-in est conçu pour fonctionner avec un ensemble spécifique de présentations qui servent de pools de diapositives qui peuvent être `sourceSlideIds` insérées. Dans ce cas, vous ou le client devez créer et gérer une source de données qui met en corrélation un critère de sélection (comme des titres ou des images) avec des ID de diapositive ou de création de diapositives qui ont été créés à partir de l’ensemble de présentations sources possibles.
