---
title: Insérer des diapositives dans une présentation PowerPoint
description: Découvrez comment insérer des diapositives d’une présentation dans une autre.
ms.date: 03/07/2021
ms.localizationpriority: medium
ms.openlocfilehash: a31933de4272634394dc6c36aafa973c41265471
ms.sourcegitcommit: 54a7dc07e5f31dd5111e4efee3e85b4643c4bef5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/21/2022
ms.locfileid: "67857570"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a>Insérer des diapositives dans une présentation PowerPoint

Un complément PowerPoint peut insérer des diapositives d’une présentation dans la présentation actuelle à l’aide de la bibliothèque JavaScript spécifique à l’application de PowerPoint. Vous pouvez contrôler si les diapositives insérées conservent la mise en forme de la présentation source ou la mise en forme de la présentation cible.

Les API d’insertion de diapositives sont principalement utilisées dans les scénarios de modèle de présentation : il existe un petit nombre de présentations connues qui servent de pools de diapositives qui peuvent être insérées par le complément. Dans un tel scénario, vous ou le client devez créer et gérer une source de données qui met en corrélation le critère de sélection (comme les titres de diapositives ou les images) avec des ID de diapositive. Les API peuvent également être utilisées dans des scénarios où l’utilisateur peut insérer des diapositives à partir d’une présentation arbitraire, mais dans ce scénario, l’utilisateur est en fait limité à l’insertion de *toutes les* diapositives à partir de la présentation source. Pour plus [d’informations à ce sujet, consultez Sélection des diapositives à insérer](#selecting-which-slides-to-insert) .

Il existe deux étapes pour insérer des diapositives d’une présentation dans une autre.

1. Convertissez le fichier de présentation source (.pptx) en chaîne au format base64.
1. Utilisez la `insertSlidesFromBase64` méthode pour insérer une ou plusieurs diapositives du fichier base64 dans la présentation actuelle.

## <a name="convert-the-source-presentation-to-base64"></a>Convertir la présentation source en base64

Il existe de nombreuses façons de convertir un fichier en base64. Le langage de programmation et la bibliothèque que vous utilisez, et s’il faut effectuer une conversion côté serveur de votre complément ou côté client, est déterminé par votre scénario. Le plus souvent, vous effectuez la conversion en JavaScript côté client à l’aide d’un objet [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) . L’exemple suivant illustre cette pratique.

1. Commencez par obtenir une référence au fichier PowerPoint source. Dans cet exemple, nous allons utiliser un `<input>` contrôle de type `file` pour inviter l’utilisateur à choisir un fichier. Ajoutez le balisage suivant à la page du complément.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    Ce balisage ajoute l’interface utilisateur dans la capture d’écran suivante à la page.

    ![Capture d’écran montrant un contrôle d’entrée de type de fichier HTML précédé d’une phrase d’instruction indiquant « Sélectionner une présentation PowerPoint à partir de laquelle insérer des diapositives ». Le contrôle se compose d’un bouton intitulé « Choisir un fichier », suivi de la phrase « Aucun fichier choisi ».](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > Il existe de nombreuses autres façons d’obtenir un fichier PowerPoint. Par exemple, si le fichier est stocké sur OneDrive ou SharePoint, vous pouvez utiliser Microsoft Graph pour le télécharger. Pour plus d’informations, consultez [Utilisation des fichiers dans Microsoft Graph](/graph/api/resources/onedrive) et [Accéder aux fichiers avec Microsoft Graph](/training/modules/msgraph-access-file-data/).

2. Ajoutez le code suivant au Code JavaScript du complément pour affecter une fonction à l’événement du `change` contrôle d’entrée. (Vous créez la `storeFileAsBase64` fonction à l’étape suivante.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Ajoutez le code suivant. Notez ce qui suit à propos de ce code.

    - La `reader.readAsDataURL` méthode convertit le fichier en base64 et le stocke dans la `reader.result` propriété. Une fois la méthode terminée, elle déclenche le gestionnaire d’événements `onload` .
    - Le `onload` gestionnaire d’événements supprime les métadonnées du fichier encodé et stocke la chaîne encodée dans une variable globale.
    - La chaîne encodée en base64 est stockée globalement, car elle sera lue par une autre fonction que vous créerez dans une étape ultérieure.

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

Votre complément insère des diapositives d’une autre présentation PowerPoint dans la présentation actuelle avec la méthode [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1)) . Voici un exemple simple dans lequel toutes les diapositives de la présentation source sont insérées au début de la présentation actuelle et où les diapositives insérées conservent la mise en forme du fichier source. Notez qu’il `chosenFileBase64` s’agit d’une variable globale qui contient une version encodée en base64 d’un fichier de présentation PowerPoint.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

Vous pouvez contrôler certains aspects du résultat d’insertion, notamment l’emplacement d’insertion des diapositives et l’obtention de la mise en forme source ou cible, en passant un objet [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) comme second paramètre à `insertSlidesFromBase64`. Voici un exemple. Tenez compte du code suivant :

- Il existe deux valeurs possibles pour la `formatting` propriété : « UseDestinationTheme » et « KeepSourceFormatting ». Si vous le souhaitez, vous pouvez utiliser l’énumération `InsertSlideFormatting` (par exemple, `PowerPoint.InsertSlideFormatting.useDestinationTheme`).
- La fonction insère les diapositives de la présentation source immédiatement après la diapositive spécifiée par la `targetSlideId` propriété. La valeur de cette propriété est une chaîne de l’une des trois formes possibles : ***nnn*#**, **#* mmmmmmmmm*** ou **_nnn_#* mmmmmmmmm***, où *nnn* est l’ID de la diapositive (généralement 3 chiffres) et *mmmmmmmmm* est l’ID de création de la diapositive (généralement 9 chiffres). Voici quelques exemples : `267#763315295`, `267#`et `#763315295`.

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

Bien sûr, vous ne saurez généralement pas au moment du codage l’ID ou l’ID de création de la diapositive cible. Plus généralement, un complément demande aux utilisateurs de sélectionner la diapositive cible. Les étapes suivantes montrent comment obtenir l’ID ***nnn*#** de la diapositive actuellement sélectionnée et comment l’utiliser comme diapositive cible.

1. Créez une fonction qui obtient l’ID de la diapositive actuellement sélectionnée à l’aide de la méthode [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) des API JavaScript courantes. Voici un exemple. Notez que l’appel est `getSelectedDataAsync` incorporé dans une fonction de retour de promesse. Pour plus d’informations sur la raison et la procédure à suivre, consultez [Wrap Common-APIs dans les fonctions de retour de promesse](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).

 
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

1. Appelez votre nouvelle fonction à l’intérieur de [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) de la fonction principale et transmettez l’ID qu’elle retourne (concaténé avec le symbole « # ») comme valeur de la `targetSlideId` propriété du `InsertSlideOptions` paramètre. Voici un exemple.

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

Vous pouvez également utiliser le paramètre [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) pour contrôler les diapositives de la présentation source qui sont insérées. Pour ce faire, affectez un tableau des ID de diapositive de la présentation source à la `sourceSlideIds` propriété. Voici un exemple qui insère quatre diapositives. Notez que chaque chaîne du tableau doit suivre l’un ou l’autre des modèles utilisés pour la `targetSlideId` propriété.

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
> Les diapositives sont insérées dans l’ordre relatif dans lequel elles apparaissent dans la présentation source, quel que soit l’ordre dans lequel elles apparaissent dans le tableau.

Il n’existe aucun moyen pratique pour les utilisateurs de découvrir l’ID ou l’ID de création d’une diapositive dans la présentation source. Pour cette raison, vous ne pouvez vraiment utiliser la `sourceSlideIds` propriété que lorsque vous connaissez les ID sources au moment du codage ou que votre complément peut les récupérer au moment de l’exécution à partir d’une source de données. Étant donné que les utilisateurs ne peuvent pas mémoriser les ID de diapositive, vous avez également besoin d’un moyen de permettre à l’utilisateur de sélectionner des diapositives, peut-être par titre ou par image, puis de mettre en corrélation chaque titre ou image avec l’ID de la diapositive.

Par conséquent, la `sourceSlideIds` propriété est principalement utilisée dans les scénarios de modèle de présentation : le complément est conçu pour fonctionner avec un ensemble spécifique de présentations qui servent de pools de diapositives pouvant être insérées. Dans un tel scénario, vous ou le client devez créer et gérer une source de données qui met en corrélation un critère de sélection (par exemple, des titres ou des images) avec des ID de diapositive ou des ID de création de diapositives qui ont été construits à partir de l’ensemble de présentations sources possibles.
