---
title: Afficher ou masquer un complément Office dans un runtime partagé
description: Découvrez comment masquer ou afficher par programme l’interface utilisateur d’un complément pendant qu’il s’exécute en continu
ms.date: 05/17/2020
localization_priority: Normal
ms.openlocfilehash: e09fa7d0a39c7157823911307558889e2ade89db
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430568"
---
# <a name="show-or-hide-an-office-add-in-in-a-shared-runtime"></a>Afficher ou masquer un complément Office dans un runtime partagé

Un complément Office peut inclure n’importe lequel des éléments suivants :

- Un volet Office
- Fichier de fonctions sans interface utilisateur (fonctions personnalisées qui n’utilisent pas de volet de tâches ou d’autres éléments d’interface utilisateur)
- Une fonction personnalisée Excel

Par défaut, chaque partie s’exécute dans son propre Runtime JavaScript distinct, avec son propre objet global et ses propres variables globales.

Il est possible que des compléments avec deux ou plusieurs composants partagent un Runtime JavaScript commun. Cette fonctionnalité d’exécution partagée permet de nouvelles API masquant et rouvrir le volet Office pendant l’exécution du complément.

## <a name="configure-an-add-in-to-use-a-shared-runtime"></a>Configurer un complément pour utiliser un runtime partagé

Pour configurer le complément afin qu’il utilise un runtime partagé, reportez-vous à la rubrique [Configure Your Office Add-in use a Shared Runtime](configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="show-and-hide-the-task-pane"></a>Afficher et masquer le volet Office

Les nouvelles API se trouvent dans la `Office.addin` propriété. Pour afficher le volet Office, votre code appelle `Office.addin.showAsTaskpane()` . Office affiche dans un volet des tâches la page que vous avez affectée à l’ID de ressource ( `resid` ) pour le volet de tâches. Il s’agit du `resid` que vous avez affecté à l' `<SourceLocation>` du `<Action xsi:type="ShowTaskpane">` dans le manifeste. (Consultez [la rubrique Configure Your Office Add-in to use a Shared Runtime](configure-your-add-in-to-use-a-shared-runtime.md).)

Il s’agit d’une méthode asynchrone, de sorte que votre code doit l’attendre lorsque le code suivant ne doit pas s’exécuter tant qu’il n’est pas terminé. Attendez la fin de cette opération avec le `await` mot clé ou une `then()` méthode, en fonction de la syntaxe JavaScript que vous utilisez. Voici un exemple de feuille de calcul Excel nommée **CurrentQuarterSales**. Le complément doit faire apparaître le volet Office chaque fois que cette feuille de calcul est activée. La méthode `onCurrentQuarter` est un gestionnaire pour l’événement [Office. Worksheet. onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) qui a été enregistré pour la feuille de calcul.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

Pour masquer le volet Office, votre code appelle `Office.addin.hide()` . L’exemple suivant est un gestionnaire inscrit pour l’événement [Office. Worksheet. onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated) .

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### <a name="preservation-of-state-and-event-listeners"></a>Conservation des écouteurs d’État et d’événement

Les `hide()` `showAsTaskpane()` méthodes et modifient uniquement la *visibilité* du volet Office. Ils ne déchargent pas ou ne le rechargent pas (ou réinitialisent son état).

Prenons le scénario suivant : un volet Office est conçu avec des onglets. L’onglet **Accueil** est ouvert lors du premier lancement du complément. Supposons qu’un utilisateur ouvre l’onglet **paramètres** et, plus tard, le code dans les appels de volet de tâches `hide()` en réponse à un événement. Toujours des appels de code plus récents `showAsTaskpane()` en réponse à un autre événement. Le volet des tâches réapparaît et l’onglet **paramètres** est toujours sélectionné.

![Capture d’écran de volet de tâches qui comporte quatre onglets intitulé Accueil, paramètres, favoris et comptes.](../images/TaskpaneWithTabs.png)

De plus, tout écouteur d’événement enregistré dans le volet Office continue de s’exécuter même si le volet Office est masqué.

Prenons le scénario suivant : le volet Office dispose d’un gestionnaire enregistré pour Excel `Worksheet.onActivated` et des `Worksheet.onDeactivated` événements pour une feuille nommée **Sheet1**. Le gestionnaire activé provoque l’affichage d’un point vert dans le volet Office. Le gestionnaire désactivé transforme le point rouge (il s’agit de son état par défaut). Supposons que le code appelle `hide()` lorsque la **feuille Sheet1** n’est pas activée et que le point est rouge. Lorsque le volet Office est masqué, la **feuille Sheet1** est activée. Appels de code ultérieurs `showAsTaskpane()` en réponse à un événement. Lorsque le volet Office s’ouvre, le point est vert car les écouteurs et gestionnaires d’événements ont été exécutés même si le volet Office a été masqué.

### <a name="handle-visibility-changed-event"></a>Événement de modification de la visibilité des handles

Lorsque votre code modifie la visibilité du volet Office avec `showAsTaskpane()` ou `hide()` , Office déclenche l' `VisibilityModeChanged` événement. Il peut être utile de gérer cet événement. Par exemple, supposons que le volet Office affiche une liste de toutes les feuilles dans un classeur. Si une nouvelle feuille de calcul est ajoutée alors que le volet Office est masqué, le fait de rendre le volet Office visible ne lui permet pas d’ajouter le nouveau nom de feuille de calcul à la liste. Toutefois, votre code peut répondre à l' `VisibilityModeChanged` événement pour recharger la propriété [Worksheet.Name](/javascript/api/excel/excel.worksheet#name) de toutes les feuilles de calcul dans la collection [Workbook. Worksheets](/javascript/api/excel/excel.workbook#worksheets) , comme illustré dans l’exemple de code ci-dessous.

Pour enregistrer un gestionnaire pour l’événement, n’utilisez pas la méthode « Add Handler » comme vous le feriez dans la plupart des contextes JavaScript Office. Au lieu de cela, il existe une fonction spéciale à laquelle vous transmettez votre gestionnaire : [Office. AddIn. onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-). Voici un exemple. Notez que la `args.visibilityMode` propriété est de type [visibilityMode](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

La fonction renvoie une autre fonction qui *annule l’enregistrement* du gestionnaire. Voici un exemple simple, mais non fiable :

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

La `onVisibilityModeChanged` méthode est asynchrone, ce qui signifie que si votre code appelle le gestionnaire d' *annulation* qui est `onVisibilityModeChanged` renvoyé, vous devez vous assurer qu’il `onVisibilityModeChanged` a été terminé avant d’appeler le gestionnaire d’annulation. Pour ce faire, vous pouvez utiliser le `await` mot clé sur l’appel de la méthode, comme dans l’exemple suivant.

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

Si vous souhaitez utiliser uniquement JavaScript pre-ES2015, votre code peut utiliser la `then` méthode pour attendre que l’objet promesse renvoyé ait été résolu et affecter la fonction renvoyée à une variable globale comme dans l’exemple suivant.

```javascript
var removeVisibilityModeHandler;

Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
}).then(function(removeHandler) {
        removeVisibilityModeHandler = removeHandler;
    });

// In some later code path, deregister with:
removeVisibilityModeHandler();
```

La fonction de désinscription est elle-même asynchrone. Par conséquent, si vous avez du code qui ne doit pas s’exécuter jusqu’à ce que la désinscription soit terminée, la fonction de désinscription doit également être attendue avec le `await` mot clé ou une `then` méthode, comme dans les exemples suivants.

Pour annuler l’inscription du gestionnaire :

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```
