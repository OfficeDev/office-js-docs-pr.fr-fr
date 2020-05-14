---
title: Exécuter du code dans votre complément Excel lorsque le document s’ouvre
description: Exécutez le code dans votre complément Excel lorsque le document s’ouvre.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: 0a9090315a4ddca80e25a94092c779a3f3271087
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217948"
---
# <a name="run-code-in-your-excel-add-in-when-the-document-opens"></a>Exécuter du code dans votre complément Excel lorsque le document s’ouvre

Vous pouvez configurer votre complément Excel de sorte qu’il charge et exécute le code dès que le document est ouvert. Cette opération est utile si vous devez enregistrer des gestionnaires d’événements, précharger des données pour le volet Office, synchroniser l’interface utilisateur ou effectuer d’autres tâches avant que le complément ne soit visible.

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Configurer votre complément pour qu’il se charge lors de l’ouverture du document

Le code suivant configure votre complément de sorte qu’il se charge et commence à s’exécuter à l’ouverture du document.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> La `setStartupBehavior` méthode est asynchrone.

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Configurer votre complément pour aucun comportement de chargement sur le document ouvert

Le code suivant permet de configurer votre complément de sorte qu’il ne démarre pas lorsque le document est ouvert. Au lieu de cela, il démarre lorsque l’utilisateur l’engage (par exemple, en choisissant un bouton de ruban ou en ouvrant le volet Office).

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Obtenir le comportement de chargement actuel

Pour déterminer le comportement de démarrage actuel, exécutez la fonction suivante, qui renvoie un objet Office. StartupBehavior.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a>Procédure d’exécution du code lorsque le document s’ouvre

Lorsque votre complément est configuré pour être chargé à l’ouverture d’un document, il s’exécutera immédiatement. Le `Office.initialize` Gestionnaire d’événements est appelé. Placez votre code de démarrage dans le `Office.initialize` Gestionnaire d’événements.

Le code suivant montre comment enregistrer un gestionnaire d’événements pour les événements de modification à partir de la feuille de calcul active. Si vous configurez le chargement de votre complément à l’ouverture du document, ce code enregistrera le gestionnaire d’événements à l’ouverture du document. Vous pouvez gérer les événements de modification avant l’ouverture du volet Office.


```JavaScript
//This is called as soon as the document opens.
//Put your startup code here.
Office.initialize = () => {
  // Add the event handler
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

## <a name="see-also"></a>Voir aussi

- [Partager des données et des événements entre des fonctions personnalisées Excel et un didacticiel de volet de tâches](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)