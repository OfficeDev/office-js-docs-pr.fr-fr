# <a name="outlook-add-in-api-requirement-set-15"></a>Ensemble de conditions requises de l’API de complément Outlook 1.5

Le sous-ensemble de l’API de complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises précédent.

## <a name="whats-new-in-15"></a>Nouveautés de la version 1.5

L’ensemble de conditions requises de la version 1.5 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). Les fonctionnalités suivantes ont été ajoutées :

- Prise en charge des [volets Office épinglables](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).
- Prise en charge de l’appel des [API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api).
- Possibilité de marquer une pièce jointe comme élément inclus.
- Ajoute de la possibilité de fermer un volet Office ou une boîte de dialogue.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback) : ajoute un gestionnaire d’événements pour un événement pris en charge.
- Ajout de [Office.EventType](office.md#eventtype-string) : spécifie l’événement associé à un gestionnaire d’événements et prend en charge l’événement ItemChanged.
- Ajout de [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string) : obtient l’URL du point de terminaison REST de ce compte de courrier.
- Modification de la méthode [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback) : cette nouvelle version comprend une nouvelle signature (`getCallbackTokenAsync([options], callback)`). La version d’origine est toujours disponible et reste inchangée.
- Ajout de [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- Modification de [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) : nouvelle valeur du dictionnaire `options` appelée `isInline`. Elle indique qu’une image est incluse dans le corps du message.
- Modification de [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incluse dans le corps du message.
- Modification de œ[Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incluse dans le corps du message.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)