---
title: "Ensemble de conditions requises de l’API du complément\_Outlook\_1.4"
description: Fonctionnalités et API introduites pour les Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1.4.
ms.date: 05/17/2021
ms.localizationpriority: medium
---

# <a name="outlook-add-in-api-requirement-set-14"></a>Ensemble de conditions requises de l’API du complément Outlook 1.4

Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-14"></a>Nouveautés de la version 1.4

L’ensemble de conditions requises 1.4 inclut toutes les fonctionnalités de l’ensemble [de conditions requises 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Il comprend en plus l’accès à l’espace de noms `Office.ui`.

### <a name="change-log"></a>Journal des modifications

- Ajout [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-displaydialogasync-member(1)) : affiche une boîte de dialogue dans Office application.
- Ajout de la méthode[Office.context.ui.messageParent](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-messageparent-member(1)): Remet un message de la part de la boîte de dialogue à sa page parent/d’ouverture.
- Ajout de l’objet [Dialog](/javascript/api/office/office.dialog?view=outlook-js-1.4&preserve-view=true): objet renvoyé lorsque la méthode [`displayDialogAsync`](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-displaydialogasync-member(1))est appelée.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
