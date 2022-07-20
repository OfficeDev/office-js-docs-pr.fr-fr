---
title: Afficher ou masquer le volet des tâches de votre complément Office
description: Découvrez comment masquer ou afficher par programmation l’interface utilisateur d’un complément pendant son exécution continue.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 76243d9e593f06eec52fe558832a722317b88c69
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889225"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Afficher ou masquer le volet des tâches de votre complément Office

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Vous pouvez afficher le volet Office de votre complément Office en appelant la `Office.addin.showAsTaskpane()` fonction.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

Le code précédent suppose un scénario dans lequel il existe une feuille de calcul Excel nommée **CurrentQuarterSales**. Le complément rend le volet Office visible chaque fois que cette feuille de calcul est activée. La méthode `onCurrentQuarter` est un gestionnaire pour l’événement [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-onactivated-member) qui a été inscrit pour la feuille de calcul.

Vous pouvez également masquer le volet Office en appelant la `Office.addin.hide()` fonction.

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

Le code précédent est un gestionnaire inscrit pour l’événement [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-ondeactivated-member) .

## <a name="additional-details-on-showing-the-task-pane"></a>Détails supplémentaires sur l’affichage du volet Office

Lorsque vous appelez `Office.addin.showAsTaskpane()`, Office affiche dans un volet Office le fichier que vous avez affecté en tant que valeur d’ID de ressource (`resid`) du volet Office. Cette `resid` valeur peut être affectée ou modifiée en ouvrant votre fichier **manifest.xml** et en se trouvant à l’intérieur de **\<SourceLocation\>** l’élément `<Action xsi:type="ShowTaskpane">` .
(Pour plus d’informations, consultez [Configurer votre complément Office pour utiliser un runtime partagé](configure-your-add-in-to-use-a-shared-runtime.md) .)

Étant donné `Office.addin.showAsTaskpane()` qu’il s’agit d’une méthode asynchrone, votre code continue à s’exécuter jusqu’à ce que la fonction soit terminée. Attendez cette fin avec le `await` mot clé ou une `then()` méthode, en fonction de la syntaxe JavaScript que vous utilisez.

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>Configurer votre complément pour utiliser le runtime partagé

Pour utiliser les méthodes et `hide()` les `showAsTaskpane()` méthodes, votre complément doit utiliser le runtime partagé. Pour plus d’informations, consultez [Configurer votre complément Office pour utiliser un runtime partagé](configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="preservation-of-state-and-event-listeners"></a>Préservation de l’état et des écouteurs d’événements

Les `hide()` méthodes et `showAsTaskpane()` les méthodes modifient uniquement la *visibilité* du volet Office. Ils ne le déchargent pas ou ne le rechargent pas (ni ne réinitialisent son état).

Prenons le scénario suivant : un volet Office est conçu avec des onglets. L’onglet **Accueil** est ouvert lors du premier lancement du complément. Supposons qu’un utilisateur ouvre l’onglet **Paramètres** et, plus tard, que le code du volet Office appelle `hide()` en réponse à un événement. Des appels de code encore ultérieurs `showAsTaskpane()` en réponse à un autre événement. Le volet Office réapparaît et l’onglet **Paramètres** est toujours sélectionné.

![Volet Office comportant quatre onglets intitulés Accueil, Paramètres, Favoris et Comptes.](../images/TaskpaneWithTabs.png)

De plus, tous les écouteurs d’événements inscrits dans le volet Office continuent à s’exécuter même lorsque le volet Office est masqué.

Prenons le scénario suivant : le volet Office comporte un gestionnaire inscrit pour Excel `Worksheet.onActivated` et `Worksheet.onDeactivated` des événements pour une feuille nommée **Sheet1**. Le gestionnaire activé provoque l’apparition d’un point vert dans le volet Office. Le gestionnaire désactivé transforme le point en rouge (qui est son état par défaut). Supposons que le code appelle `hide()` lorsque **sheet1** n’est pas activé et que le point est rouge. Bien que le volet Office soit masqué, **la feuille Sheet1** est activée. Appels `showAsTaskpane()` de code ultérieurs en réponse à un événement. Lorsque le volet Office s’ouvre, le point est vert, car les écouteurs et gestionnaires d’événements ont été exécutés même si le volet Office était masqué.

## <a name="handle-the-visibility-changed-event"></a>Gérer l’événement de changement de visibilité

Lorsque votre code modifie la visibilité du volet Office avec `showAsTaskpane()` ou `hide()`, Office déclenche l’événement `VisibilityModeChanged` . Il peut être utile de gérer cet événement. Par exemple, supposons que le volet Office affiche une liste de toutes les feuilles d’un classeur. Si une nouvelle feuille de calcul est ajoutée pendant que le volet Office est masqué, rendre le volet Office visible n’ajouterait pas en soi le nouveau nom de feuille de calcul à la liste. Toutefois, votre code peut répondre à l’événement `VisibilityModeChanged` pour recharger la propriété [Worksheet.name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member) de toutes les feuilles de calcul de la collection [Workbook.worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member) , comme indiqué dans l’exemple de code ci-dessous.

Pour inscrire un gestionnaire pour l’événement, vous n’utilisez pas de méthode « ajouter un gestionnaire » comme vous le feriez dans la plupart des contextes JavaScript Office. Au lieu de cela, il existe une fonction spéciale à laquelle vous passez votre gestionnaire : [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#office-office-addin-onvisibilitymodechanged-member(1)). Voici un exemple. Notez que la `args.visibilityMode` propriété est de type [VisibilityMode](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

La fonction retourne une autre fonction qui *désinscrit* le gestionnaire. Voici un exemple simple, mais pas robuste.

```javascript
const removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

La `onVisibilityModeChanged` méthode est asynchrone et retourne une promesse, ce qui signifie que votre code doit attendre la réalisation de la promesse avant de pouvoir appeler le gestionnaire **de désinscription** .

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
const removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

La fonction de désinscription est également asynchrone et retourne une promesse. Par conséquent, si vous avez du code qui ne doit pas s’exécuter tant que la désinscription n’est pas terminée, vous devez attendre la promesse retournée par la fonction de désinscription.

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>Voir aussi

- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](configure-your-add-in-to-use-a-shared-runtime.md)
- [Exécuter un cote dans votre complément Office lors de l’ouverture du document](run-code-on-document-open.md)
