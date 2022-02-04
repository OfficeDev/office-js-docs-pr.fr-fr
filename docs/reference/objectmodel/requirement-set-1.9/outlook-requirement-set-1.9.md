---
title: Outlook l’ensemble de conditions requises de l’API du add-in 1.9
description: Ensemble de conditions requises 1.9 pour Outlook API de votre application.
ms.date: 05/17/2021
ms.localizationpriority: medium
---

# <a name="outlook-add-in-api-requirement-set-19"></a>Outlook l’ensemble de conditions requises de l’API du add-in 1.9

Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-19"></a>Nouveautés de la 1.9

L’ensemble de conditions requises 1.9 inclut toutes les fonctionnalités de l’ensemble [de conditions requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). Les fonctionnalités suivantes ont été ajoutées.

- Ajout de nouvelles API pour l’ajout à l’envoi, les propriétés personnalisées et les fonctionnalités de formulaire d’affichage.
- Prise en charge supplémentaire pour `Dialog.messageChild`.

### <a name="change-log"></a>Journal des modifications

- Ajout de [CustomProperties.getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#outlook-office-customproperties-getall-member(1)) : ajoute une nouvelle fonction `CustomProperties` à l’objet qui obtient toutes les propriétés personnalisées.
- Ajout de [Dialog.messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) : ajoute une nouvelle méthode qui fournit un message à partir de la page hôte, tel qu’un volet Des tâches ou un fichier de fonction sans interface utilisateur, à une boîte de dialogue qui a été ouverte à partir de la page.
- Ajout de [l’élément manifeste ExtendedPermissions](../../manifest/extendedpermissions.md) : ajoute un élément enfant à [l’élément manifeste VersionOverrides](../../manifest/versionoverrides.md) . Pour qu’un module de prise en charge de la fonctionnalité d’ajout à l’envoi[,](../../../outlook/append-on-send.md)`AppendOnSend` l’autorisation étendue doit être incluse dans la collection d’autorisations étendues.
- Ajout [Office.context.mailbox.displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displayappointmentformasync-member(1))`Mailbox` : ajoute une nouvelle fonction à l’objet qui affiche un rendez-vous existant. Il s’agit de la version async de la `displayAppointmentForm` méthode.
- Ajout [Office.context.mailbox.displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaymessageformasync-member(1))`Mailbox` : ajoute une nouvelle fonction à l’objet qui affiche un message existant. Il s’agit de la version async de la `displayMessageForm` méthode.
- Ajout [Office.context.mailbox.displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaynewappointmentformasync-member(1))`Mailbox` : ajoute une nouvelle fonction à l’objet qui affiche un nouveau formulaire de rendez-vous. Il s’agit de la version async de la `displayNewAppointmentForm` méthode.
- Ajout [Office.context.mailbox.displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaynewmessageformasync-member(1))`Mailbox` : ajoute une nouvelle fonction à l’objet qui affiche un nouveau formulaire de message. Il s’agit de la version async de la `displayNewMessageForm` méthode.
- Ajout de [Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#outlook-office-body-appendonsendasync-member(1))`Body` : ajoute une nouvelle fonction à l’objet qui ajoute des données à la fin du corps de l’élément en mode composition.
- Ajout [Office.context.mailbox.item.displayReplyAllFormAsync](office.context.mailbox.item.md#methods)`Item` : ajoute une nouvelle fonction à l’objet qui affiche le formulaire « Répondre à tous » en mode lecture. Il s’agit de la version async de la `displayReplyAllForm` méthode.
- Ajout [Office.context.mailbox.item.displayReplyFormAsync](office.context.mailbox.item.md#methods)`Item` : ajoute une nouvelle fonction à l’objet qui affiche le formulaire « Répondre » en mode lecture. Il s’agit de la version async de la `displayReplyForm` méthode.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
