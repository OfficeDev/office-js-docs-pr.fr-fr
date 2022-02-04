---
title: Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook
description: "Obtenez ou définissez diverses propriétés d’un élément dans un complément\_Outlook d’un scénario de composition, y compris ses destinataires, son objet, son corps, et ses emplacement et heure de rendez-vous."
ms.date: 12/10/2019
ms.localizationpriority: medium
---

# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook

Découvrez comment obtenir ou définir diverses propriétés d’un élément dans un complément Outlook d’un scénario de composition, y compris ses destinataires, son objet, son corps, et ses emplacement et heure de rendez-vous.

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>Obtention et définition des propriétés d’un élément pour un complément de composition

Dans un formulaire de composition, vous pouvez obtenir la plupart des propriétés qui sont exposées sur le même genre d’élément que dans un formulaire de lecture (comme attendees, recipients, subject et body), et vous pouvez obtenir quelques propriétés supplémentaires qui sont pertinentes uniquement dans un formulaire de composition mais pas dans un formulaire de lecture (body, bcc).

Pour la plupart de ces propriétés, comme il est possible qu’un complément Outlook et l’utilisateur modifient la même propriété dans l’interface utilisateur en même temps, les méthodes d’obtention et de définition de ces propriétés sont asynchrones. Le tableau 1 énumère les propriétés de niveau élément et les méthodes asynchrones correspondantes pour les obtenir et les définir dans un formulaire de composition. Les propriétés  [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [item.conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) constituent des exceptions, car les utilisateurs ne peuvent pas les modifier. Vous pouvez les obtenir par programmation de la même façon dans un formulaire de composition et dans un formulaire de lecture, directement à partir de l’objet parent.

Outre l’accès aux propriétés d’élément dans l’API JavaScript Office, vous pouvez accéder aux propriétés au niveau de l’élément à l’aide Exchange Web Services (EWS). Avec l’autorisation **ReadWriteMailbox**, vous pouvez utiliser la méthode [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour accéder aux opérations EWS, [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) and [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation), pour obtenir et définir plus de propriétés d’au moins un élément dans la boîte aux lettres de l’utilisateur.

La fonction `makeEwsRequestAsync` est disponible à la fois dans les formulaires de lecture et de composition. Pour plus d’informations sur l’autorisation **ReadWriteMailbox** et l’accès à EWS par le biais de la plateforme des Compléments Office, consultez les rubriques [Présentation des autorisations de complément Outlook](understanding-outlook-add-in-permissions.md) et [Appeler des services Web à partir d’un complément Outlook](web-services.md).

**Tableau 1. Méthodes asynchrones pour obtenir ou définir des propriétés d’élément dans un formulaire de composition**

<br/>

| Propriété | Type de propriété | Méthode asynchrone d’obtention | Méthode(s) asynchrone(s) de définition |
|:-----|:-----|:-----|:-----|
|[bbc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Recipients](/javascript/api/outlook/office.recipients)|[Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))|[Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))|
|[body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Corps](/javascript/api/outlook/office.body)|[Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))|[Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)), [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))|
|[cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Time](/javascript/api/outlook/office.time)|[Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1))|[Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))|
|[location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Location](/javascript/api/outlook/office.location)|[Location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1))|[Location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))|
|[optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Heure|Time.getAsync|Time.setAsync|
|[subject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Subject](/javascript/api/outlook/office.subject)|[Subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1))|[Subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))|
|[to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## <a name="see-also"></a>Voir aussi

- [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md)
- [Présentation des autorisations de complément Outlook](understanding-outlook-add-in-permissions.md)
- [Appeler des services web à partir d’un complément Outlook](web-services.md)
- [Obtention et définition de données d’élément Outlook dans des formulaires de lecture ou de composition](item-data.md)
