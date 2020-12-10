---
title: Insertion et suppression de diapositives dans une présentation PowerPoint
description: Découvrez comment insérer des diapositives d’une présentation dans une autre et comment supprimer des diapositives.
ms.date: 12/04/2020
localization_priority: Normal
ms.openlocfilehash: ceb78054a95ac4b26bd71f79a086a00e3dce5278
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49613703"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation-preview"></a>Insertion et suppression de diapositives dans une présentation PowerPoint (aperçu)

Un complément PowerPoint peut insérer des diapositives d’une présentation dans la présentation en cours à l’aide de la bibliothèque JavaScript propre à l’application de PowerPoint. Vous pouvez contrôler si les diapositives insérées conservent la mise en forme de la présentation source ou la mise en forme de la présentation cible. Vous pouvez également supprimer des diapositives de la présentation.

[!include[General preview API prerequisites](../includes/using-preview-apis-host.md)]

Les API d’insertion de diapositives sont principalement utilisées dans les scénarios de modèle de présentation : il existe un petit nombre de présentations connues qui servent de pools de diapositives qui peuvent être insérées par le complément. Dans ce cas, vous ou le client devez créer et gérer une source de données qui met en corrélation le critère de sélection (comme les titres ou les images des diapositives) et les ID de diapositive. Les API peuvent également être utilisées dans les scénarios dans lesquels l’utilisateur peut insérer des diapositives à partir de n’importe quelle présentation arbitraire, mais dans ce scénario, l’utilisateur est limité à insérer *toutes* les diapositives de la présentation source. Pour plus d’informations à ce sujet, voir [sélection des diapositives à insérer](#selecting-which-slides-to-insert) .

Il existe deux étapes pour insérer des diapositives d’une présentation dans une autre.

1. Convertissez le fichier de présentation source (. pptx) en une chaîne au format Base64.
1. Utilisez la `insertSlidesFromBase64` méthode pour insérer une ou plusieurs diapositives à partir du fichier Base64 dans la présentation active.

## <a name="convert-the-source-presentation-to-base64"></a>Convertir la présentation source en base64

Il existe plusieurs façons de convertir un fichier en base64. Le langage de programmation et la bibliothèque que vous utilisez et s’il faut effectuer une conversion côté serveur de votre complément ou côté client est déterminé par votre scénario. En règle générale, vous effectuerez la conversion en JavaScript côté client à l’aide d’un objet [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) . L’exemple suivant illustre cette pratique.

1. Commencez par obtenir une référence au fichier PowerPoint source. Dans cet exemple, nous allons utiliser un `<input>` contrôle de type `file` pour inviter l’utilisateur à choisir un fichier. Ajoutez le balisage suivant à la page de complément.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    Ce balisage ajoute l’interface utilisateur dans la capture d’écran suivante à la page :

    ![Capture d’écran illustrant un contrôle d’entrée de type de fichier HTML précédé d’une phrase pédagogique en lisant « sélectionnez une présentation PowerPoint à partir de laquelle insérer des diapositives ». Le contrôle se compose d’un bouton intitulé « choisir le fichier » suivi de la phrase « aucun fichier choisi ».](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > Il existe de nombreuses autres façons d’obtenir un fichier PowerPoint. Par exemple, si le fichier est stocké sur OneDrive ou SharePoint, vous pouvez utiliser Microsoft Graph pour le télécharger. Pour plus d’informations, consultez la rubrique [utilisation de fichiers dans Microsoft Graph](/graph/api/resources/onedrive) et [accès à des fichiers avec Microsoft Graph](/learn/modules/msgraph-access-file-data/).

2. Ajoutez le code suivant au JavaScript du complément pour assigner une fonction à l’événement du contrôle d’entrée `change` . (Vous créez la `storeFileAsBase64` fonction à l’étape suivante.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Ajoutez le code suivant. Notez ce qui suit à propos de ce code :

    - La `reader.readAsDataURL` méthode convertit le fichier en base64 et le stocke dans la `reader.result` propriété. Une fois la méthode terminée, le gestionnaire d’événements est déclenché `onload` .
    - Le `onload` Gestionnaire d’événements supprime les métadonnées du fichier encodé et stocke la chaîne encodée dans une variable globale.
    - La chaîne codée en base64 est stockée globalement, car elle sera lue par une autre fonction que vous créez dans une étape ultérieure.

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

Votre complément insère des diapositives d’une autre présentation PowerPoint dans la présentation actuelle à l’aide de la méthode [Presentation. insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) . Voici un exemple simple dans lequel toutes les diapositives de la présentation source sont insérées au début de la présentation en cours et les diapositives insérées conservent la mise en forme du fichier source. Notez qu' `chosenFileBase64` il s’agit d’une variable globale qui contient une version codée en base64 d’un fichier de présentation PowerPoint.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

Vous pouvez contrôler certains aspects du résultat de l’insertion, y compris où les diapositives sont insérées et déterminer si elles obtiennent la mise en forme source ou cible, en transmettant un objet [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) comme deuxième paramètre à `insertSlidesFromBase64` . Voici un exemple. Tenez compte du code suivant :

- Il existe deux valeurs possibles pour la `formatting` propriété : « UseDestinationTheme » et « KeepSourceFormatting ». Vous pouvez également utiliser l' `InsertSlideFormatting` énumération (par exemple, `PowerPoint.InsertSlideFormatting.useDestinationTheme` ).
- La fonction insère les diapositives de la présentation source immédiatement après la diapositive spécifiée par la `targetSlideId` propriété. La valeur de cette propriété est une chaîne d’une des trois formes possibles : ***nnn * #**, * *#* mmmmmmmmm * * * ou **_nnn_ #* mmmmmmmmm * * *, où *nnn* est l’ID de la diapositive (généralement 3 chiffres) et *mmmmmmmmm* est l’ID de création de la diapositive (généralement 9 chiffres). Voici quelques exemples :, `267#763315295` `267#` et `#763315295` .

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

Bien entendu, vous ne saurez généralement pas au moment du code l’ID ou l’ID de création de la diapositive cible. Plus communément, un complément demande aux utilisateurs de sélectionner la diapositive cible. Les étapes suivantes montrent comment obtenir l’ID ***nnn * #** de la diapositive actuellement sélectionnée et l’utiliser comme diapositive cible.

1. Créez une fonction qui obtient l’ID de la diapositive actuellement sélectionnée à l’aide de la [Office.context.docméthode ument. getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) des API JavaScript communes. Voici un exemple. Notez que l’appel à `getSelectedDataAsync` est incorporé dans une fonction de retour à la vente. Pour plus d’informations sur les raisons et la procédure à suivre, consultez [la rubrique Wrap Common-APIs dans les fonctions de retour à la vente](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).

 
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

1. Appelez votre nouvelle fonction à l’intérieur de [PowerPoint. Run ()](/javascript/api/powerpoint#PowerPoint_run_batch_) de la fonction main et transmettez l’ID qu’elle renvoie (concaténé avec le symbole « # ») en tant que valeur de la `targetSlideId` propriété du `InsertSlideOptions` paramètre. Voici un exemple.

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

Vous pouvez également utiliser le paramètre [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) pour contrôler les diapositives de la présentation source qui doivent être insérées. Pour ce faire, affectez un tableau des ID de diapositives de la présentation source à la `sourceSlideIds` propriété. Voici un exemple qui insère quatre diapositives. Notez que chaque chaîne dans le tableau doit respecter un ou l’autre des modèles utilisés pour la `targetSlideId` propriété.

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
> Les diapositives sont insérées dans le même ordre relatif que celui dans lequel elles apparaissent dans la présentation source, quel que soit l’ordre dans lequel elles apparaissent dans le tableau.

Il n’existe aucun moyen pratique pour les utilisateurs de découvrir l’ID ou l’ID de création d’une diapositive dans la présentation source. Pour cette raison, vous pouvez uniquement utiliser la `sourceSlideIds` propriété lorsque vous avez identifié les ID de source au moment du codage ou que votre complément peut les récupérer lors de l’exécution à partir d’une source de données. Étant donné que les utilisateurs ne peuvent pas mémoriser les ID de diapositive, vous avez également besoin d’un moyen pour permettre à l’utilisateur de sélectionner des diapositives, par exemple par titre ou par image, puis de corréler chaque titre ou image avec l’ID de la diapositive.

En conséquence, la `sourceSlideIds` propriété est principalement utilisée dans les scénarios de modèle de présentation : le complément est conçu pour fonctionner avec un ensemble spécifique de présentations qui servent de pools de diapositives qui peuvent être insérées. Dans ce cas, vous ou le client devez créer et gérer une source de données qui met en corrélation un critère de sélection (tel que des titres ou des images) avec des ID de diapositive ou des ID de création de diapositives qui ont été créés à partir de l’ensemble de présentations source possibles.

## <a name="delete-slides"></a>Supprimer des diapositives

Vous pouvez supprimer une diapositive en obtenant une référence à l’objet [Slide](/javascript/api/powerpoint/powerpoint.slide) qui représente la diapositive et appeler la `Slide.delete` méthode. Voici un exemple dans lequel la quatrième diapositive est supprimée.

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
