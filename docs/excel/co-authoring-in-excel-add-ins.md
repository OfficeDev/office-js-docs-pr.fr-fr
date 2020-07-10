---
title: Co-création dans des macros complémentaires Excel
description: Apprenez à co-auteur d’un classeur Excel stocké dans OneDrive, OneDrive entreprise ou SharePoint Online.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 4414bf64f05c29328c63d0857a6e498495712ff1
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093475"
---
# <a name="coauthoring-in-excel-add-ins"></a>Co-création dans des macros complémentaires Excel  

With [coauthoring](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104), multiple people can work together and edit the same Excel workbook simultaneously. All coauthors of a workbook can see another coauthor's changes as soon as that coauthor saves the workbook. To coauthor an Excel workbook, the workbook must be stored in OneDrive, OneDrive for Business, or SharePoint Online.

> [!IMPORTANT]
> Dans Excel pour Microsoft 365, vous remarquerez l’enregistrement automatique dans le coin supérieur gauche. Lorsque l’enregistrement automatique est activé, les co-auteurs visualisent vos modifications en temps réel. Prenez en considération l’impact de ce comportement sur la conception de votre complément Excel. Les utilisateurs peuvent désactiver l’enregistrement automatique via le commutateur dans le coin supérieur gauche de la fenêtre Excel.

## <a name="coauthoring-overview"></a>Vue d’ensemble de la co-création

Lorsque vous modifiez le contenu d’un classeur, Excel synchronise automatiquement ces modifications avec tous les co-auteurs. Les co-auteurs peuvent modifier le contenu d’un classeur, mais peuvent également exécuter du code dans un complément Excel. Par exemple, lorsque le code JavaScript suivant s’exécute dans un complément Office, la valeur de la plage est définie sur Contoso:

```js
range.values = [['Contoso']];
```
Après la synchronisation de « Contoso » avec tous les co-auteurs, tout utilisateur ou complément en cours d’exécution dans le même classeur visualisera la nouvelle valeur de la plage.

Coauthoring only synchronizes the content within the shared workbook. Values copied from the workbook to JavaScript variables in an Excel add-in are not synchronized. For example, if your add-in stores the value of a cell (such as 'Contoso') in a JavaScript variable, and then a coauthor changes the value of the cell to 'Example', after synchronization all coauthors see 'Example' in the cell. However, the value of the JavaScript variable is still set to 'Contoso'. Furthermore, when multiple coauthors use the same add-in, each coauthor has their own copy of the variable, which is not synchronized. When you use variables that use workbook content, be sure you check for updated values in the workbook before you use the variable.

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>Utiliser des événements pour gérer l’état de la mémoire de votre complément

Excel add-ins can read workbook content (from hidden worksheets and a setting object), and then store it in data structures such as variables. After the original values are copied into any of these data structures, coauthors can update the original workbook content. This means that the copied values in the data structures are now out of sync with the workbook content. When you build your add-ins, be sure to account for this separation of workbook content and values stored in data structures.

For example, you might build a content add-in that displays custom visualizations. The state of your custom visualizations might be saved in a hidden worksheet. When coauthors use the same workbook, the following scenario can occur:

- User A opens the document and the custom visualizations are shown in the workbook. The custom visualizations read data from a hidden worksheet (for example, the color of the visualizations is set to blue).
- User B opens the same document, and starts modifying the custom visualizations. User B sets the color of the custom visualizations to orange. Orange is saved to the hidden worksheet.
- La feuille de calcul masquée de l’utilisateur A est mise à jour avec la nouvelle valeur Orange.
- Les visualisations personnalisées de l’utilisateur A sont toujours bleues.

If you want User A's custom visualizations to respond to changes made by coauthors on the hidden worksheet, use the [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event. This ensures that changes to workbook content made by coauthors is reflected in the state of your add-in.

## <a name="caveats-to-using-events-with-coauthoring"></a>Restrictions à l’utilisation des événements dans le cadre de la co-création

As described earlier, in some scenarios, triggering events for all coauthors provides an improved user experience. However, be aware that in some scenarios this behavior can produce poor user experiences. 

Par exemple, dans les scénarios de validation de données, il est fréquent d’afficher l’interface utilisateur en réponse aux événements. L’événement [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) décrit dans la section précédente s’exécute lorsqu’un utilisateur local ou un co-auteur modifie (à distance) le contenu du classeur dans la liaison. Si le gestionnaire d’événements de l' `BindingDataChanged` événement affiche l’interface utilisateur, les utilisateurs voient l’interface utilisateur qui n’est pas liée aux modifications sur lesquelles ils travaillaient dans le classeur, ce qui entraîne une expérience utilisateur médiocre. Évitez d’afficher l’interface utilisateur lorsque vous utilisez des événements dans votre complément.

## <a name="see-also"></a>Voir aussi

- [À propos de la co-création dans Excel (VBA)](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [Impact de l’enregistrement automatique sur les compléments et les macros (VBA)](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
