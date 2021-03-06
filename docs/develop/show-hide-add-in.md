---
title: Afficher ou masquer le volet Office de votre add-in Office
description: Découvrez comment masquer ou afficher par programme l’interface utilisateur d’un add-in pendant qu’il s’exécute en continu.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 20db609a3a6ded5624391f705dab1ad6b8f6e043
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789229"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Afficher ou masquer le volet Office de votre add-in Office

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Vous pouvez afficher le volet Office de votre add-in Office en appelant la `Office.addin.showAsTaskpane()` fonction.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

Le code précédent suppose un scénario dans lequel il existe une feuille de calcul Excel nommée **CurrentQuarterSales**. Le add-in rend le volet Des tâches visible chaque fois que cette feuille de calcul est activée. La méthode est un handler pour `onCurrentQuarter` l’événement [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) qui a été inscrit pour la feuille de calcul.

Vous pouvez également masquer le volet Des tâches en appelant la `Office.addin.hide()` fonction.

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

Le code précédent est un handler inscrit pour [l’événement Office.Worksheet.onDeactivated.](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated)

## <a name="additional-details-on-showing-the-task-pane"></a>Détails supplémentaires sur l’affichage du volet Des tâches

Lorsque vous appelez, Office affiche dans un volet Office le fichier que vous avez affecté en tant qu’ID de ressource ( ) du `Office.addin.showAsTaskpane()` `resid` volet Office. Cette valeur peut être affectée ou modifiée en ouvrant votre fichiermanifest.xmlet en `resid` le localisant à  `<SourceLocation>` l’intérieur de `<Action xsi:type="ShowTaskpane">` l’élément.
(Pour plus [d’informations,](configure-your-add-in-to-use-a-shared-runtime.md) voir Configurer votre complément Office pour utiliser un runtime partagé.)

Étant `Office.addin.showAsTaskpane()` donné qu’il s’agit d’une méthode asynchrone, votre code continuera à s’exécute jusqu’à ce que la fonction soit terminée. Attendez cette fin avec le mot clé ou une méthode, en fonction de la `await` `then()` syntaxe JavaScript que vous utilisez.

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>Configurer votre add-in pour utiliser le runtime partagé

Pour utiliser les `showAsTaskpane()` méthodes et les `hide()` méthodes, votre add-in doit utiliser le runtime partagé. Pour plus d’informations, voir [Configurer votre add-in Office pour utiliser un runtime partagé.](configure-your-add-in-to-use-a-shared-runtime.md)

## <a name="preservation-of-state-and-event-listeners"></a>Conservation des écouteurs d’état et d’événements

Les `hide()` méthodes et les méthodes `showAsTaskpane()` modifient uniquement la *visibilité* du volet Des tâches. Ils ne déchargent pas ou ne rechargent pas (ou réinitialisent son état).

Envisagez le scénario suivant : un volet Des tâches est conçu avec des onglets. **L’onglet** Accueil est ouvert lors du premier lancement du module. Supposons qu’un utilisateur ouvre l’onglet **Paramètres** et, plus tard, code dans les appels du volet Des tâches en `hide()` réponse à un événement. Appels de code ultérieurs `showAsTaskpane()` en réponse à un autre événement. Le volet Des tâches réapparaît et l’onglet **Paramètres** est toujours sélectionné.

![Capture d’écran du volet Des tâches avec quatre onglets étiquetés Accueil, Paramètres, Favoris et Comptes.](../images/TaskpaneWithTabs.png)

En outre, tous les écouteurs d’événements inscrits dans le volet Des tâches continuent de s’exécuter même lorsque le volet Des tâches est masqué.

Envisagez le scénario suivant : le volet Des tâches possède un handler inscrit pour Excel et des événements pour une `Worksheet.onActivated` `Worksheet.onDeactivated` feuille nommée **Sheet1**. Le handler activé entraîne l’apparition d’un point vert dans le volet Des tâches. Le handler désactivé transforme le point en rouge (qui est son état par défaut). Supposons alors que le code appelle `hide()` **lorsque la feuille Sheet1 n’est** pas activée et que le point est rouge. Bien que le volet Des tâches soit masqué, **la feuille Sheet1** est activée. Appels de code `showAsTaskpane()` ultérieurs en réponse à un événement. Lorsque le volet Des tâches s’ouvre, le point est vert, car les écouteurs et les handlers d’événements s’ouvrent même si le volet Des tâches a été masqué.

## <a name="handle-the-visibility-changed-event"></a>Gérer l’événement de changement de visibilité

Lorsque votre code modifie la visibilité du volet Office avec `showAsTaskpane()` `hide()` ou, Office déclenche `VisibilityModeChanged` l’événement. Il peut être utile de gérer cet événement. Par exemple, supposons que le volet Des tâches affiche une liste de toutes les feuilles dans un workbook. Si une nouvelle feuille de calcul est ajoutée alors que le volet Des tâches est masqué, le fait de rendre le volet Des tâches visible n’ajoute pas en soi le nouveau nom de feuille de calcul à la liste. Toutefois, votre code peut répondre à l’événement pour recharger la propriété Worksheet.name de toutes les feuilles de calcul de la `VisibilityModeChanged` collection [](/javascript/api/excel/excel.worksheet#name) [Workbook.worksheets,](/javascript/api/excel/excel.workbook#worksheets) comme illustré dans l’exemple de code ci-dessous.

Pour inscrire un handler pour l’événement, vous n’utilisez pas de méthode « add handler » comme vous le feriez dans la plupart des contextes JavaScript Office. Au lieu de cela, il existe une fonction spéciale à laquelle vous passez votre handler : [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-). Voici un exemple. Notez que `args.visibilityMode` la propriété est de type [VisibilityMode](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

La fonction renvoie une autre fonction *qui désinsère* le handler. Voici un exemple simple, mais non robuste :

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

La `onVisibilityModeChanged` méthode est asynchrone et renvoie une promesse, ce qui signifie que votre code  doit attendre la réalisation de la promesse avant de pouvoir appeler le sous-enregistré.

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

La fonction d’agrégation est également asynchrone et renvoie une promesse. Ainsi, si vous avez du code qui ne doit pas s’exécuter tant que l’agrégation n’est pas terminée, vous devez attendre la promesse renvoyée par la fonction d’agrégation.

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>Voir aussi

- [Configurer votre add-in Office pour utiliser un runtime JavaScript partagé](configure-your-add-in-to-use-a-shared-runtime.md)
- [Exécuter du code dans votre add-in Office à l’ouverture du document](run-code-on-document-open.md)
