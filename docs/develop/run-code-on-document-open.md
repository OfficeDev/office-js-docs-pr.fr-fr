---
title: Exécuter un cote dans votre complément Office lors de l’ouverture du document
description: Découvrez comment exécuter du code dans votre complément Office lorsque le document s’ouvre.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: ef580151a5b3289c801f3e872988cbb3474bd8e0
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422916"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>Exécuter un cote dans votre complément Office lors de l’ouverture du document

[!include[Shared runtime requirements](../includes/shared-runtime-requirements-note.md)]

Vous pouvez configurer votre complément Office pour charger et exécuter du code dès que le document est ouvert. Cela est utile si vous devez inscrire des gestionnaires d’événements, précharger des données pour le volet Office, synchroniser l’interface utilisateur ou effectuer d’autres tâches avant que le complément ne soit visible.

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Configurer votre complément pour qu’il se charge à l’ouverture du document

Le code suivant configure votre complément pour qu’il charge et commence à s’exécuter lorsque le document est ouvert.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> La `setStartupBehavior` méthode est asynchrone.

## <a name="place-startup-code-in-officeinitialize"></a>Placer le code de démarrage dans Office.initialize

Lorsque votre complément est configuré pour être chargé sur le document ouvert, il s’exécute immédiatement. Le `Office.initialize` gestionnaire d’événements sera appelé. Placez votre code de démarrage dans le ou `Office.onReady` le gestionnaire d’événements`Office.initialize`.

Le code de complément Excel suivant montre comment inscrire un gestionnaire d’événements pour les événements de modification à partir de la feuille de calcul active. Si vous configurez votre complément pour qu’il se charge sur le document ouvert, ce code inscrit le gestionnaire d’événements lors de l’ouverture du document. Vous pouvez gérer les événements de modification avant l’ouverture du volet Office.

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
    await Excel.run(async (context) => {    
        await context.sync();
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);
  });
}
```

Le code de complément PowerPoint suivant montre comment inscrire un gestionnaire d’événements pour les événements de modification de sélection à partir du document PowerPoint. Si vous configurez votre complément pour qu’il se charge sur le document ouvert, ce code inscrit le gestionnaire d’événements lors de l’ouverture du document. Vous pouvez gérer les événements de modification avant l’ouverture du volet Office.

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

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Configurer votre complément pour aucun comportement de chargement lors de l’ouverture du document

Le code suivant configure votre complément pour qu’il ne démarre pas lorsque le document est ouvert. Au lieu de cela, il démarre lorsque l’utilisateur l’engage d’une manière ou d’une autre, par exemple en choisissant un bouton de ruban ou en ouvrant le volet Office.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Obtenir le comportement de charge actuel

Pour déterminer le comportement de démarrage actuel, exécutez la méthode suivante, qui retourne un `Office.StartupBehavior` objet.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="see-also"></a>Voir aussi

- [Configurer votre complément Office pour utiliser un runtime partagé](configure-your-add-in-to-use-a-shared-runtime.md)
- [Partager des données et des événements entre les fonctions personnalisées Excel et le didacticiel du volet Office](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Utilisation d’événements à l’aide de l’API JavaScript pour Excel](../excel/excel-add-ins-events.md)
- [Runtimes dans les compléments Office](../testing/runtimes.md)
