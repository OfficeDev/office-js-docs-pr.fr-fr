---
title: Co-création dans des macros complémentaires Excel
description: Apprenez à co-Excel un Excel stocké dans OneDrive, OneDrive Entreprise ou SharePoint Online.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7fd2e2846c4256e7aac1ffa7263b4aa57b744d21
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744383"
---
# <a name="coauthoring-in-excel-add-ins"></a>Co-création dans des macros complémentaires Excel  

Avec la [co-création](https://support.microsoft.com/office/7152aa8b-b791-414c-a3bb-3024e46fb104), plusieurs personnes peuvent travailler ensemble et modifier simultanément le même classeur Excel. Tous les co-auteurs d’un classeur peuvent voir les modifications d’un autre co-auteur dès que ce co-auteur enregistre le classeur. Pour co-créer un classeur Excel, le classeur doit être enregistré dans OneDrive, OneDrive Entreprise ou SharePoint Online.

> [!IMPORTANT]
> Dans Excel pour Microsoft 365, vous remarquerez AutoSave dans le coin supérieur gauche. Lorsque l’enregistrement automatique est activé, les co-auteurs visualisent vos modifications en temps réel. Prenez en considération l’impact de ce comportement sur la conception de votre complément Excel. Les utilisateurs peuvent désactiver l’enregistrement automatique via le commutateur dans le coin supérieur gauche de la fenêtre Excel.

## <a name="coauthoring-overview"></a>Vue d’ensemble de la co-création

Lorsque vous modifiez le contenu d’un classeur, Excel synchronise automatiquement ces modifications avec tous les co-auteurs. Les co-auteurs peuvent modifier le contenu d’un classeur, mais peuvent également exécuter du code dans un complément Excel. Par exemple, lorsque le code JavaScript suivant s’exécute dans un Office, la valeur d’une plage est définie sur Contoso.

```js
range.values = [['Contoso']];
```

Après la synchronisation de « Contoso » avec tous les co-auteurs, tout utilisateur ou complément en cours d’exécution dans le même classeur visualisera la nouvelle valeur de la plage.

La co-création permet uniquement la synchronisation du contenu dans le classeur partagé. Les valeurs copiées du classeur vers les variables JavaScript dans un complément Excel ne sont pas synchronisées. Par exemple, si votre complément enregistre la valeur d’une cellule (par exemple, « Contoso ») dans une variable JavaScript et qu’un co-auteur modifie ensuite la valeur de la cellule sur « Exemple », après la synchronisation, tous les co-auteurs verront « Exemple » dans la cellule. Toutefois, la valeur de la variable JavaScript sera toujours définie sur « Contoso ». En outre, lorsque plusieurs co-auteurs utilisent le même complément, chaque co-auteur possède sa propre copie de la variable, qui n’est pas synchronisée. Lorsque vous utilisez des variables qui utilisent le contenu du classeur, veillez à bien rechercher les valeurs mises à jour dans le classeur avant d’utiliser la variable.

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>Utiliser des événements pour gérer l’état de la mémoire de votre complément

Les compléments Excel peuvent lire le contenu du classeur (à partir de feuilles de calcul masquées et d’un objet de paramètres), puis l’enregistrer dans des structures de données comme des variables. Une fois que les valeurs d’origine sont copiées dans l’une de ces structures de données, les co-auteurs peuvent mettre à jour le contenu du classeur d’origine. Cela signifie que les valeurs copiées dans les structures de données ne sont plus synchronisées avec le contenu du classeur. Lorsque vous générez vos compléments, pensez à prendre en compte cette séparation de contenu du classeur et les valeurs enregistrées dans les structures de données.

Par exemple, vous pouvez créer un complément de contenu qui affiche des visualisations personnalisées. L’état de vos visualisations personnalisées peut être enregistré dans une feuille de calcul masquée. Lorsque les co-auteurs utilisent le même workbook, le scénario suivant peut se produire.

- L’utilisateur A ouvre le document et les visualisations personnalisées sont affichées dans le classeur. Les visualisations personnalisées lisent les données d’une feuille de calcul masquée (par exemple, la couleur des visualisations est définie sur bleu).
- L’utilisateur B ouvre le même document et commence à modifier les visualisations personnalisées. L’utilisateur B définit la couleur des visualisations personnalisées sur orange. La valeur Orange est enregistrée dans la feuille de calcul masquée.
- La feuille de calcul masquée de l’utilisateur A est mise à jour avec la nouvelle valeur Orange.
- Les visualisations personnalisées de l’utilisateur A sont toujours bleues.

Si vous souhaitez que les visualisations personnalisées de l’utilisateur A répondent aux modifications apportées par les co-auteurs sur la feuille de calcul masquée, utilisez l’événement [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs). Cela garantit que les modifications apportées au contenu du classeur par les co-auteurs sont répercutées à l’état de votre complément.

## <a name="caveats-to-using-events-with-coauthoring"></a>Restrictions à l’utilisation des événements dans le cadre de la co-création

Comme indiqué précédemment, dans certains scénarios, le déclenchement d’événements pour tous les co-auteurs permet d’améliorer l’expérience utilisateur. Toutefois, sachez que, dans certains scénarios, ce comportement peut entraîner des expériences utilisateur médiocres.

Par exemple, dans les scénarios de validation de données, il est fréquent d’afficher l’interface utilisateur en réponse aux événements. L’événement [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) décrit dans la section précédente s’exécute lorsqu’un utilisateur local ou un co-auteur modifie (à distance) le contenu du classeur dans la liaison. Si le handler `BindingDataChanged` d’événement de l’événement affiche l’interface utilisateur, les utilisateurs voient une interface utilisateur qui n’est pas liée aux modifications qu’ils utilisaient dans le classeur, entraînant une expérience utilisateur médiocre. Évitez d’afficher l’interface utilisateur lorsque vous utilisez des événements dans votre complément.

## <a name="avoid-table-row-coauthoring-conflicts"></a>Éviter les conflits de co-auteur de lignes de tableau

Il s’agit d’un problème connu : les appels à l’API [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)) peuvent provoquer des conflits de co-édition. Nous vous déconseillons d’utiliser cette API si vous prévoyez d’exécuter votre application pendant que d’autres utilisateurs modifient le workbook du module (en particulier, s’ils modifient le tableau ou une plage sous le tableau). Les instructions suivantes doivent vous aider à éviter les problèmes avec la méthode (et éviter le déclenchement de la barre jaune Excel affiche qui demande aux utilisateurs `TableRowCollection.add` d’actualiser).

1. Utilisez [`Range.values`](/javascript/api/excel/excel.range#excel-excel-range-values-member) au lieu de [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)). La définition `Range` des valeurs directement sous le tableau développe automatiquement le tableau. Sinon, l’ajout de lignes de tableau via les `Table` API entraîne des conflits de fusion pour les utilisateurs coauth.
1. Aucune règle de [validation](https://support.microsoft.com/office/29fecbcc-d1b9-42c1-9d76-eff3ce5f7249) des données ne doit être appliquée aux cellules sous le tableau, sauf si la validation des données est appliquée à la colonne entière.
1. S’il existe des données sous le tableau, le add-in doit le gérer avant de définir la valeur de la plage. L’insertion [`Range.insert`](/javascript/api/excel/excel.range#excel-excel-range-insert-member(1)) d’une ligne vide déplace les données et fait de l’espace pour le tableau en développement. Sinon, vous risquez de overwriting cells below the table.
1. Vous ne pouvez pas ajouter une ligne vide à un tableau avec `Range.values`. Le tableau se développe automatiquement uniquement si des données sont présentes dans les cellules directement en dessous du tableau. Utilisez des données temporaires ou des colonnes masquées comme solution de contournement pour ajouter une ligne de tableau vide.

## <a name="see-also"></a>Voir aussi

- [À propos de la co-création dans Excel (VBA)](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [Impact de l’enregistrement automatique sur les compléments et les macros (VBA)](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
