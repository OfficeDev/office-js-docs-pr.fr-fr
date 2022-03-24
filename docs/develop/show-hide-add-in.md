---
title: Afficher ou masquer le volet des tâches de votre complément Office
description: Découvrez comment masquer ou afficher par programme l’interface utilisateur d’un add-in pendant qu’il s’exécute en continu.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7e881f5fc0d5258aa886709a0aee2eee5836feef
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743958"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Afficher ou masquer le volet des tâches de votre complément Office

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Vous pouvez afficher le volet Des tâches de votre Office en appelant la `Office.addin.showAsTaskpane()` fonction.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

Le code précédent suppose un scénario dans lequel il existe une feuille de Excel nommée **CurrentQuarterSales**. Le add-in rend le volet Des tâches visible chaque fois que cette feuille de calcul est activée. La méthode `onCurrentQuarter` est un handler pour le [Office. Événement Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-onactivated-member) qui a été inscrit pour la feuille de calcul.

Vous pouvez également masquer le volet Des tâches en appelant la `Office.addin.hide()` fonction.

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

Le code précédent est un handler inscrit pour le [Office. Événement Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-ondeactivated-member).

## <a name="additional-details-on-showing-the-task-pane"></a>Détails supplémentaires sur l’affichage du volet Des tâches

Lorsque vous appelez `Office.addin.showAsTaskpane()`, Office affiche dans un volet Des tâches le fichier que vous avez affecté en tant que valeur d’ID de ressource (`resid`) du volet De tâches. Cette `resid` valeur peut être affectée ou modifiée en ouvrant votre fichier **manifest.xml** et en le localisant `<SourceLocation>` à l’intérieur de l’élément `<Action xsi:type="ShowTaskpane">` .
(Pour [plus d’informations, voir Configurer Office complément](configure-your-add-in-to-use-a-shared-runtime.md) pour utiliser un runtime partagé.)

Étant `Office.addin.showAsTaskpane()` donné qu’il s’agit d’une méthode asynchrone, votre code continue d’être en cours d’exécution jusqu’à ce que la fonction soit terminée. Attendez cette fin avec le mot `await` clé ou `then()` une méthode, en fonction de la syntaxe JavaScript que vous utilisez.

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>Configurer votre add-in pour utiliser le runtime partagé

Pour utiliser les méthodes `showAsTaskpane()` et les `hide()` méthodes, votre add-in doit utiliser le runtime partagé. Pour plus d’informations, voir [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="preservation-of-state-and-event-listeners"></a>Conservation des écouteurs d’état et d’événements

Les `hide()` méthodes `showAsTaskpane()` et les méthodes modifient uniquement *la visibilité* du volet Des tâches. Ils ne le déchargent pas, ne le rechargent pas (ou ne réinitialisent pas son état).

Envisagez le scénario suivant : un volet Des tâches est conçu avec des onglets. **L’onglet** Accueil est ouvert lors du premier lancement du module. Supposons qu’un utilisateur ouvre **l’onglet Paramètres** et, plus tard, le code `hide()` dans les appels du volet Des tâches en réponse à un événement. Appels de code ultérieurs en `showAsTaskpane()` réponse à un autre événement. Le volet Des tâches réapparaît et **l’onglet Paramètres** est toujours sélectionné.

![Capture d’écran du volet Des tâches avec quatre onglets étiquetés Accueil, Paramètres, Favoris et Comptes.](../images/TaskpaneWithTabs.png)

En outre, tous les écouteurs d’événements inscrits dans le volet Des tâches continuent de s’exécuter même lorsque le volet Des tâches est masqué.

Envisagez le scénario suivant : le volet Des tâches possède un handler `Worksheet.onActivated` `Worksheet.onDeactivated` inscrit pour les Excel et les événements d’une feuille nommée **Sheet1**. Le handler activé entraîne l’apparition d’un point vert dans le volet Des tâches. Le handler désactivé transforme le point en rouge (qui est son état par défaut). Supposons alors que le code appelle `hide()` **lorsque la feuille Sheet1 n’est** pas activée et que le point est rouge. Bien que le volet Des tâches soit masqué, **la feuille Sheet1** est activée. Appels de code ultérieurs `showAsTaskpane()` en réponse à un événement. Lorsque le volet Des tâches s’ouvre, le point est vert, car les écouteurs et les handlers d’événements s’ouvrent même si le volet Des tâches a été masqué.

## <a name="handle-the-visibility-changed-event"></a>Gérer l’événement de changement de visibilité

Lorsque votre code modifie la visibilité du volet Des `showAsTaskpane()` `hide()`tâches avec ou, Office déclenche l’événement`VisibilityModeChanged`. Il peut être utile de gérer cet événement. Par exemple, supposons que le volet Des tâches affiche une liste de toutes les feuilles d’un workbook. Si une nouvelle feuille de calcul est ajoutée alors que le volet Des tâches est masqué, le fait de rendre le volet Des tâches visible n’ajoute pas en soi le nouveau nom de feuille de calcul à la liste. Toutefois, votre code `VisibilityModeChanged` peut répondre à l’événement pour recharger la propriété [Worksheet.name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member) de toutes les feuilles de calcul de la collection [Workbook.worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member) , comme illustré dans l’exemple de code ci-dessous.

Pour inscrire un handler pour l’événement, vous n’utilisez pas de méthode « add handler » comme vous le feriez dans la plupart Office contextes JavaScript. Au lieu de cela, il existe une fonction spéciale à laquelle vous passez votre [Office:Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#office-office-addin-onvisibilitymodechanged-member(1)). Voici un exemple. Notez que la `args.visibilityMode` propriété est de type [VisibilityMode](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

La fonction renvoie une autre fonction *qui désinsère* le handler. Voici un exemple simple, mais non robuste.

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

La `onVisibilityModeChanged` méthode est asynchrone et renvoie une promesse, ce qui signifie que votre code doit attendre la fin de la promesse avant de pouvoir appeler le sous-enregistré.

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

- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](configure-your-add-in-to-use-a-shared-runtime.md)
- [Exécuter un cote dans votre complément Office lors de l’ouverture du document](run-code-on-document-open.md)
