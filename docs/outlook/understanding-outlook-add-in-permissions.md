---
title: Présentation des autorisations de complément Outlook
description: Les compléments Outlook spécifient le niveau d’autorisation requis dans leur manifeste (Restricted, ReadItem, ReadWriteItem ou ReadWriteMailbox).
ms.date: 02/19/2020
ms.localizationpriority: medium
ms.openlocfilehash: b515ef470331a513d6b57007f372b3e4dec1d25b
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660227"
---
# <a name="understanding-outlook-add-in-permissions"></a>Présentation des autorisations de complément Outlook

Les compléments Outlook spécifient le niveau d’autorisation requis dans leur manifeste. Les niveaux disponibles sont **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox**. Ces niveaux d’autorisation sont cumulatifs : **Restricted** est le niveau le plus bas, et chaque niveau supérieur inclut les autorisations de tous les niveaux inférieurs. **ReadWriteMailbox** contient toutes les autorisations prises en charge.

Vous pouvez voir les autorisations demandées par un complément de messagerie avant de l’installer depuis [AppSource](https://appsource.microsoft.com). Vous pouvez également voir les autorisations requises des compléments installés dans le Centre d’administration Exchange.

## <a name="restricted-permission"></a>Autorisation Restricted

L’autorisation **Restricted** est la plus basique. Indiquez **Restricted** dans l’élément [Permissions](/javascript/api/manifest/permissions) du manifeste pour demander cette autorisation. Outlook affecte par défaut ce niveau d’autorisation à un complément de messagerie si le complément ne demande pas d’autorisation spécifique dans son manifeste.

### <a name="can-do"></a>Vous pouvez :

- [Obtenir uniquement des entités spécifiques](match-strings-in-an-item-as-well-known-entities.md) (numéro de téléphone, adresse, URL) de l’objet ou du corps de l’élément.

- Spécifier une [règle d’activation ItemIs](activation-rules.md#itemis-rule) qui exige que l’élément actuel soit un type d’élément spécifique dans un formulaire de lecture ou de composition, ou une [règle ItemHasKnownEntity](match-strings-in-an-item-as-well-known-entities.md) qui correspond à l’un des sous-ensembles plus petits d’entités connues prises en charge (numéro de téléphone, adresse, URL) dans l’élément sélectionné.

- Accéder aux propriétés et aux méthodes qui ne sont **pas** associées aux informations spécifiques concernant l’utilisateur ou l’élément. (Consultez la section suivante pour obtenir la liste des membres qui le sont.)

### <a name="cant-do"></a>Vous ne pouvez pas :

- Utilisez une règle [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) sur le contact, l’adresse e-mail, la suggestion de réunion ou l’entité de suggestion de tâche.

- Utiliser la règle [ItemHasAttachment](/javascript/api/manifest/rule#itemhasattachment-rule) ou [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule).

- Accéder aux membres de la liste suivante qui se rapportent aux informations de l’utilisateur ou de l’élément. Si vous tentez d’accéder aux membres de cette liste, vous obtenez la valeur **null** et un message d’erreur indiquant qu’Outlook requiert le complément de messagerie pour bénéficier d’autorisations élevées.

  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.from](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.organizer](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.sender](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.userProfile](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
  - [Body](/javascript/api/outlook/office.body) et tous ses membres enfants
  - [Location](/javascript/api/outlook/office.location) et tous ses membres enfants
  - [Recipients](/javascript/api/outlook/office.recipients) et tous ses membres enfants
  - [Subject](/javascript/api/outlook/office.subject) et tous ses membres enfants
  - [Time](/javascript/api/outlook/office.time) et tous ses membres enfants

## <a name="readitem-permission"></a>Autorisation ReadItem

L’autorisation **ReadItem** est le niveau suivant dans le modèle d’autorisations. Spécifiez **ReadItem** dans l’élément **\<Permissions\>** du manifeste pour demander cette autorisation.

### <a name="can-do"></a>Vous pouvez :

- [Lire toutes les propriétés](item-data.md) de l’élément actuel dans un formulaire de lecture ou de [composition](get-and-set-item-data-in-a-compose-form.md), par exemple, [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) dans un formulaire de lecture et [item.to.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)) dans un formulaire de composition.

- [Obtenir un jeton de rappel pour obtenir les pièces jointes de l’élément](get-attachments-of-an-outlook-item.md) ou l’élément complet avec les services Web Exchange ou les [API REST Outlook](use-rest-api.md).

- [Écrire des propriétés personnalisées](/javascript/api/outlook/office.customproperties) définies par le complément sur cet élément.

- [Obtenir toutes les entités existantes connues](match-strings-in-an-item-as-well-known-entities.md), et pas seulement un sous-ensemble, à partir de l’objet ou du corps de l’élément.

- Utiliser toutes les [entités connues](activation-rules.md#itemhasknownentity-rule) dans les règles [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) ou les [expressions régulières](activation-rules.md#itemhasregularexpressionmatch-rule) dans les règles [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule). L’exemple suivant suit le schéma version 1.1. Il affiche une règle qui active le complément si une ou plusieurs des entités connues se trouvent dans l’objet ou le corps du message sélectionné.

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

- Utilisez l’une des API suivantes.
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.bcc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.bcc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1))
  - [item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))
  - [item.body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))
  - [item.cc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.cc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.end.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))
  - [item.optionalAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.optionalAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.requiredAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.start.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))
  - [item.to.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.to.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))

## <a name="readwriteitem-permission"></a>Autorisation ReadWriteItem

Spécifiez **ReadWriteItem** dans l’élément **\<Permissions\>** du manifeste pour demander cette autorisation. Les compléments de messagerie activés dans des formulaires de composition et utilisant des méthodes d’écriture (par exemple, **Message.to.addAsync** ou **Message.to.setAsync**) doivent utiliser au moins ce niveau d’autorisation.

### <a name="can-do"></a>Vous pouvez :

- [Lire et écrire toutes les propriétés au niveau de l’élément](item-data.md) concernant l’élément affiché ou en cours de composition dans Outlook.

- [Ajouter ou supprimer des pièces jointes](add-and-remove-attachments-to-an-item-in-a-compose-form.md) de cet élément.

- Utilisez tous les autres membres de l’API JavaScript Office applicables aux compléments de messagerie, à l’exception de **Mailbox.makeEWSRequestAsync**.

### <a name="cant-do"></a>Vous ne pouvez pas :

- Utilisez le jeton fourni par **mailbox.getCallbackTokenAsync** pour les actions suivantes :
  - Mettre à jour ou supprimer l’élément actuel à l’aide de l’API REST Outlook ou accéder à tous les autres éléments de la boîte aux lettres de l’utilisateur
  - Récupérer l’élément d’événement de calendrier actuel à l’aide de l’API REST Outlook

- Utiliser **mailbox.makeEWSRequestAsync**.

## <a name="readwritemailbox-permission"></a>Autorisation ReadWriteMailbox

L’autorisation **ReadWriteMailbox** correspond au plus haut niveau d’autorisation. Spécifiez **ReadWriteMailbox** dans l’élément **\<Permissions\>** du manifeste pour demander cette autorisation.

En plus des actions prises en charge par l’autorisation **ReadWriteItem**, le jeton fourni par **mailbox.getCallbackTokenAsync** fournit un accès aux opérations des services web Exchange ou à l’API REST Outlook pour effectuer les opérations suivantes :

- Lire et écrire toutes les propriétés d’un élément de la boîte aux lettres de l’utilisateur.
- Créer, lire et écrire dans tous les dossiers ou tous les éléments de cette boîte aux lettres.
- Envoyer un élément depuis cette boîte aux lettres.

Par le biais de **mailbox.makeEWSRequestAsync**, vous pouvez accéder aux opérations EWS suivantes.

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

- [Confidentialité, autorisations et sécurité pour les compléments Outlook](../concepts/privacy-and-security.md)
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md)
