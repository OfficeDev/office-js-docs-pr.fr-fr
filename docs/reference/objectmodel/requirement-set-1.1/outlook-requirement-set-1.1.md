# <a name="outlook-add-in-api-requirement-set-11"></a>Ensemble de conditions requises de l’API du complément Outlook 1.1

Le sous-ensemble de l’API pour complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente. 

## <a name="whats-new-in-11"></a>Nouveautés de la version 1.1

L’ensemble de conditions requises de la version 1.1 comprend toutes les fonctionnalités de l’ensemble de conditions requises de la version 1.0. Désormais, les compléments peuvent accéder au corps des messages et des rendez-vous et vous pouvez modifier l’élément actif.

### <a name="change-log"></a>Journal des modifications

- Ajout de l’objet [Body](/javascript/api/outlook_1_1/office.body) : fournit des méthodes pour ajouter et mettre à jour le contenu d’un élément dans un complément Outlook.
- Ajout de l’objet [Location](/javascript/api/outlook_1_1/office.location) : fournit des méthodes pour obtenir et définir le lieu d’une réunion dans un complément Outlook.
- Ajout de l’objet [Recipients](/javascript/api/outlook_1_1/office.recipients) : fournit des méthodes pour obtenir et définir les destinataires d’un rendez-vous ou d’un message dans un complément Outlook.
- Ajout de l’objet [Subject](/javascript/api/outlook_1_1/office.subject) : fournit des méthodes pour obtenir et définir l’objet d’un rendez-vous ou d’un message dans un complément Outlook.
- Ajout de l’objet [Time](/javascript/api/outlook_1_1/office.time) : fournit des méthodes pour obtenir et définir l’heure de début ou de fin d’une réunion dans un complément Outlook.
- Ajout de la méthode [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) : ajoute un fichier à un message ou un rendez-vous en pièce jointe.
- Ajout de la méthode [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback) : ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.
- Ajout de la méthode [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) : supprime une pièce jointe d’un message ou d’un rendez-vous.
- Ajout de l’objet [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-bodyjavascriptapioutlook11officebody) : obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.
- Ajout de l’objet [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipientsjavascriptapioutlook11officerecipients) : obtient ou définit les destinataires en Cci (copie carbone invisible) d’un message.
- Ajout de l’énumération [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype) : spécifie le type de destinataire d’un rendez-vous.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)