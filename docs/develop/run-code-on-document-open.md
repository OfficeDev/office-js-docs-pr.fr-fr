---
title: Exécuter un cote dans votre complément Office lors de l’ouverture du document
description: Découvrez comment exécuter du code dans votre Office de votre add-in à l’ouverture du document.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 1655c053a4fa6f92aae95f2155991fa4f7f7a5a7
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937643"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>Exécuter un cote dans votre complément Office lors de l’ouverture du document

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Vous pouvez configurer votre Office pour charger et exécuter du code dès que le document est ouvert. Cela est utile si vous devez inscrire des handlers d’événements, pré-charger des données pour le volet Des tâches, synchroniser l’interface utilisateur ou effectuer d’autres tâches avant que le module ne soit visible.

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Configurer votre add-in pour qu’il se charge à l’ouverture du document

Le code suivant configure votre add-in pour qu’il se charge et démarre l’exécution à l’ouverture du document.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> La `setStartupBehavior` méthode est asynchrone.

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Configurer votre add-in pour qu’il n’y a aucun comportement de chargement à l’ouverture d’un document

Le code suivant configure votre add-in pour qu’il ne démarre pas lorsque le document est ouvert. Au lieu de cela, il démarre lorsque l’utilisateur l’engage d’une manière ou d’une autre, par exemple en choisissant un bouton de ruban ou en ouvrant le volet Des tâches.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Obtenir le comportement de chargement actuel

Pour déterminer le comportement de démarrage actuel, exécutez la fonction suivante, qui renvoie un `Office.StartupBehavior` objet.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a>Comment exécuter du code à l’ouverture du document

Lorsque votre add-in est configuré pour se charger à l’ouverture du document, il s’exécute immédiatement. Le `Office.initialize` handler d’événements est appelé. Placez votre code de démarrage dans le `Office.initialize` ou le `Office.onReady` handler d’événements.

Le code Excel de la feuille de calcul active montre comment inscrire un handler d’événements pour les événements de modification à partir de la feuille de calcul active. Si vous configurez votre add-in pour qu’il se charge sur le document ouvert, ce code enregistre le handler d’événements lors de l’ouverture du document. Vous pouvez gérer les événements de modification avant l’ouverture du volet Des tâches.

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.initialize = () => {
  // Add the event handler.
  Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(onChange);

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
  });
};

/**
 * Handle the changed event from the worksheet.
 *
 * @param event The event information from Excel
 */
async function onChange(event) {
  return Excel.run(function(context) {
    return context.sync().then(function() {
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
    });
  });
}
```

Le code de PowerPoint suivant montre comment inscrire un handler d’événements pour les événements de modification de sélection à partir PowerPoint document. Si vous configurez votre add-in pour qu’il se charge sur le document ouvert, ce code enregistre le handler d’événements lors de l’ouverture du document. Vous pouvez gérer les événements de modification avant l’ouverture du volet Des tâches.

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
    console.log("A handler has been registered for the onChanged event.");
  }
});

/**
 * Handle the changed event from the PowerPoint document.
 *
 * @param event The event information from PowerPoint
 */
async function onChange(event) {
  console.log("Change type of event: " + event.type);
}
```

## <a name="see-also"></a>Voir aussi

- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](configure-your-add-in-to-use-a-shared-runtime.md)
- [Partager des données et des événements entre Excel fonctions personnalisées et didacticiel du volet Des tâches](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Utilisation d’événements à l’aide de l’API JavaScript pour Excel](../excel/excel-add-ins-events.md)
