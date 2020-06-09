---
title: Présentation des autorisations de complément Outlook
description: Les compléments Outlook spécifient le niveau d’autorisation requis dans leur manifeste (Restricted, ReadItem, ReadWriteItem ou ReadWriteMailbox).
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: d566c0f330ff4473fc9c71a7dff26048093707db
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611623"
---
# <a name="understanding-outlook-add-in-permissions"></a>Présentation des autorisations de complément Outlook

Les compléments Outlook spécifient le niveau d’autorisation requis dans leur manifeste. Les niveaux disponibles sont **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox**. Ces niveaux d’autorisation sont cumulatifs : **Restricted** est le niveau le plus bas, et chaque niveau supérieur inclut les autorisations de tous les niveaux inférieurs. **ReadWriteMailbox** contient toutes les autorisations prises en charge.

Vous pouvez voir les autorisations demandées par un complément de messagerie avant de l’installer depuis [AppSource](https://appsource.microsoft.com). Vous pouvez également voir les autorisations requises des compléments installés dans le Centre d’administration Exchange.

## <a name="restricted-permission"></a>Autorisation Restricted

L’autorisation **Restricted** est la plus basique. Indiquez **Restricted** dans l’élément [Permissions](../reference/manifest/permissions.md) du manifeste pour demander cette autorisation. Outlook affecte par défaut ce niveau d’autorisation à un complément de messagerie si le complément ne demande pas d’autorisation spécifique dans son manifeste.

### <a name="can-do"></a>Vous pouvez :

- [Obtenir uniquement des entités spécifiques](match-strings-in-an-item-as-well-known-entities.md) (numéro de téléphone, adresse, URL) de l’objet ou du corps de l’élément.

- Spécifier une [règle d’activation ItemIs](activation-rules.md#itemis-rule) qui exige que l’élément actuel soit un type d’élément spécifique dans un formulaire de lecture ou de composition, ou une [règle ItemHasKnownEntity](match-strings-in-an-item-as-well-known-entities.md) qui correspond à l’un des sous-ensembles plus petits d’entités connues prises en charge (numéro de téléphone, adresse, URL) dans l’élément sélectionné.

- Accéder aux propriétés et aux méthodes qui ne sont **pas** associées aux informations spécifiques concernant l’utilisateur ou l’élément. (Consultez la section suivante pour obtenir la liste des membres qui le sont.)

### <a name="cant-do"></a>Vous ne pouvez pas :

- Utiliser une règle [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) sur l’entité contact, adresse de messagerie, suggestion de réunion ou suggestion de tâche.

- Utiliser la règle [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) ou [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule).

- Accéder aux membres de la liste suivante qui se rapportent aux informations de l’utilisateur ou de l’élément. Si vous tentez d’accéder aux membres de cette liste, vous obtenez la valeur **null** et un message d’erreur indiquant qu’Outlook requiert le complément de messagerie pour bénéficier d’autorisations élevées.

    - [item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.from](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.organizer](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.sender](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.userProfile](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
    - [Body](/javascript/api/outlook/office.body) et tous ses membres enfants
    - [Location](/javascript/api/outlook/office.location) et tous ses membres enfants
    - [Recipients](/javascript/api/outlook/office.recipients) et tous ses membres enfants
    - [Subject](/javascript/api/outlook/office.subject) et tous ses membres enfants
    - [Time](/javascript/api/outlook/office.time) et tous ses membres enfants

## <a name="readitem-permission"></a>Autorisation ReadItem

L’autorisation **ReadItem** est le niveau suivant dans le modèle d’autorisations. Indiquez **ReadItem** dans l’élément **Permissions** du manifeste pour demander cette autorisation.

### <a name="can-do"></a>Vous pouvez :

- [Lire toutes les propriétés](item-data.md) de l’élément actuel dans un formulaire de lecture ou de [composition](get-and-set-item-data-in-a-compose-form.md), par exemple, [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) dans un formulaire de lecture et [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) dans un formulaire de composition.

- [Obtenir un jeton de rappel pour obtenir les pièces jointes de l’élément](get-attachments-of-an-outlook-item.md) ou l’élément complet avec les services Web Exchange ou les [API REST Outlook](use-rest-api.md).

- [Écrire des propriétés personnalisées](/javascript/api/outlook/office.CustomProperties) définies par le complément sur cet élément.

- [Obtenir toutes les entités existantes connues](match-strings-in-an-item-as-well-known-entities.md), et pas seulement un sous-ensemble, à partir de l’objet ou du corps de l’élément.

- Utiliser toutes les [entités connues](activation-rules.md#itemhasknownentity-rule) dans les règles [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ou les [expressions régulières](activation-rules.md#itemhasregularexpressionmatch-rule) dans les règles [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule). L’exemple suivant suit le schéma version 1.1. Il montre une règle qui active le complément si une ou plusieurs entités connues sont trouvées dans l’objet ou le corps du message sélectionné :

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### <a name="cant-do"></a>Vous ne pouvez pas :

- Utilisez le jeton fourni par **mailbox.getCallbackTokenAsync** pour les actions suivantes :
    - Mettre à jour ou supprimer l’élément actuel à l’aide de l’API REST Outlook ou accéder à tous les autres éléments de la boîte aux lettres de l’utilisateur
    - Récupérer l’élément d’événement de calendrier actuel à l’aide de l’API REST Outlook

- Utilisez l’une des API suivantes :
    - [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.bcc.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.bcc.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)
    - [item.body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-)
    - [item.body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)
    - [item.cc.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.cc.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.end.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [item.location.setAsync](/javascript/api/outlook/office.Location#setasync-location--options--callback-)
    - [item.optionalAttendees.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.optionalAttendees.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.requiredAttendees.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.requiredAttendees.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.start.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [item.subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)
    - [item.to.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.to.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)

## <a name="readwriteitem-permission"></a>Autorisation ReadWriteItem

Vous pouvez indiquer **ReadWriteItem** dans l’élément **Permissions** du manifeste pour demander cette autorisation. Les compléments de messagerie activés dans des formulaires de composition et utilisant des méthodes d’écriture (par exemple, **Message.to.addAsync** ou **Message.to.setAsync**) doivent utiliser au moins ce niveau d’autorisation.

### <a name="can-do"></a>Vous pouvez :

- [Lire et écrire toutes les propriétés au niveau de l’élément](item-data.md) concernant l’élément affiché ou en cours de composition dans Outlook.

- [Ajouter ou supprimer des pièces jointes](add-and-remove-attachments-to-an-item-in-a-compose-form.md) de cet élément.

- Utilisez tous les autres membres de l’API JavaScript pour Office qui s’appliquent aux compléments de messagerie, sauf **Mailbox. makeEWSRequestAsync**.

### <a name="cant-do"></a>Vous ne pouvez pas :

- Utilisez le jeton fourni par **mailbox.getCallbackTokenAsync** pour les actions suivantes :
    - Mettre à jour ou supprimer l’élément actuel à l’aide de l’API REST Outlook ou accéder à tous les autres éléments de la boîte aux lettres de l’utilisateur
    - Récupérer l’élément d’événement de calendrier actuel à l’aide de l’API REST Outlook

- Utiliser **mailbox.makeEWSRequestAsync**.

## <a name="readwritemailbox-permission"></a>Autorisation ReadWriteMailbox

L’autorisation **ReadWriteMailbox** correspond au plus haut niveau d’autorisation. Indiquez **ReadWriteMailbox** dans l’élément **Permissions** du manifeste pour demander cette autorisation.

En plus des actions prises en charge par l’autorisation **ReadWriteItem**, le jeton fourni par **mailbox.getCallbackTokenAsync** fournit un accès aux opérations des services web Exchange ou à l’API REST Outlook pour effectuer les opérations suivantes :

- Lire et écrire toutes les propriétés d’un élément de la boîte aux lettres de l’utilisateur.
- Créer, lire et écrire dans tous les dossiers ou tous les éléments de cette boîte aux lettres.
- Envoyer un élément depuis cette boîte aux lettres.

Grâce à **mailbox.makeEWSRequestAsync**, vous pouvez accéder aux opérations des services web Exchange suivantes :

- [CopyItem](/exchange/client-developer/web-service-reference/copyitem-operation)
- [CreateFolder](/exchange/client-developer/web-service-reference/createfolder-operation)
- [CreateItem](/exchange/client-developer/web-service-reference/createitem-operation)
- [FindConversation](/exchange/client-developer/web-service-reference/findconversation-operation)
- [FindFolder](/exchange/client-developer/web-service-reference/findfolder-operation)
- [FindItem](/exchange/client-developer/web-service-reference/finditem-operation)
- [GetConversationItems](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [GetFolder](/exchange/client-developer/web-service-reference/getfolder-operation)
- [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)
- [MarkAsJunk](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [MoveItem](/exchange/client-developer/web-service-reference/moveitem-operation)
- [SendItem](/exchange/client-developer/web-service-reference/senditem-operation)
- [UpdateFolder](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)

Toute tentative d’utilisation d’une opération non prise en charge entraînera une réponse d’erreur.

## <a name="see-also"></a>Voir aussi

- [Confidentialité, autorisations et sécurité pour les compléments Outlook](../develop/privacy-and-security.md)
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md)
