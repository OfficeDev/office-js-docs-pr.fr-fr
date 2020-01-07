---
title: Ensemble de conditions requises de l’API du complément Outlook 1.8
description: ''
ms.date: 10/31/2019
localization_priority: Priority
ms.openlocfilehash: 1e1420bd355c16941c7cb4ce66ecdca56e1c8927
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902153"
---
# <a name="outlook-add-in-api-requirement-set-18"></a>Ensemble de conditions requises de l’API du complément Outlook 1.8

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

## <a name="whats-new-in-18"></a>Nouveautés de la version 1.8

L’ensemble de conditions requises de la version 1.8 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md). Les fonctionnalités suivantes ont été ajoutées.

- Ajout de nouvelles fonctionnalités d’API pour les pièces jointes, de catégories, d’accès délégué, d’emplacement amélioré, d’en-têtes Internet et de blocage d’envoi.
- Ajout d’un paramètre `options` facultatif à Event.completed.
- Ajout d’une prise en charge des événements AttachmentsChanged et EnhancedLocationsChanged.

### <a name="change-log"></a>Journal des modifications

- Ajout d’[AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8) : ajoute un nouvel objet représentant le contenu d’une pièce jointe.
- Ajout de [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8) : ajoute un nouvel objet représentant les catégories d’un élément.
- Ajout de [CategoryDetails](/javascript/api/outlook/office.categorydetails?view=outlook-js-1.8) : ajoute un nouvel objet représentant les détails d’une catégorie (son nom et la couleur associée).
- Ajout d’[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) : ajoute un nouvel objet représentant l’ensemble des lieux pour un rendez-vous.
- Ajout d’[InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8) : ajout d’un nouvel objet représentant les en-têtes Internet d’un élément de message. Mode composition uniquement.
- Ajout de [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) : ajoute un nouvel objet représentant un lieu. En lecture seule.
- Ajout de [LocationIdentifier](/javascript/api/outlook/office.locationidentifier?view=outlook-js-1.8) : ajoute un nouvel objet représentant l’ID d’un lieu.
- Ajout de [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8) : ajoute un nouvel objet représentant la liste principale des catégories d’une boîte aux lettres.
- Ajout de [SharedProperties](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) : ajoute un nouvel objet représentant les propriétés d’un élément de rendez-vous ou de message dans un dossier, un calendrier ou une boîte aux lettres partagé(e).
- Ajout d’un [élément de manifeste SupportsSharedFolders](../../manifest/supportssharedfolders.md) : ajoute un élément enfant à l’élément de manifeste [DesktopFormFactor](../../manifest/desktopformfactor.md). Définit si le complément est disponible dans les scénarios de délégué.
- Ajout d’[Office.context.mailbox.masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#mastercategories) : ajoute une nouvelle propriété représentant la liste principale des catégories d’une boîte aux lettres.
- Ajout d’[Office.context.mailbox.item.categories](/javascript/api/outlook/office.item?view=outlook-js-1.8#categories) : ajoute une nouvelle propriété représentant l’ensemble des catégories d’une boîte aux lettres.
- Ajout d’[Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) : ajoute une nouvelle méthode qui vous permet de joindre un fichier à un message ou à un rendez-vous. Ce fichier est représenté par une chaîne encodée en base 64.
- Ajout d’[Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocation) : ajoute une nouvelle propriété représentant l’ensemble des lieux pour un rendez-vous.
- Ajout d'[Office.context.mailbox.item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread?view=outlook-js-1.8#getallinternetheadersasync-options--callback-) : ajoute une nouvelle méthode récupérant tous les en-têtes Internet pour un élément de message. Mode Lecture uniquement.
- Ajout d’[Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) : ajoute une nouvelle méthode pour obtenir le contenu d’une pièce jointe spécifique.
- Ajout d’[Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails) : ajoute une nouvelle méthode récupérant les pièces jointes à un élément en mode composition.
- Ajout d’[Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#getitemidasyncoptions-callback) : ajoute une nouvelle méthode obtenant l’ID d’un rendez-vous ou d’un élément de message enregistré.
- Ajout d’[Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) : ajout d’une nouvelle méthode obtenant un objet représentant les sharedProperties d’un rendez-vous ou d’un élément de message.
- Ajout d’[Office.context.mailbox.item.internetHeaders](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.8#internetheaders) : ajoute une nouvelle propriété représentant les en-têtes Internet personnalisés d’élément de message. Mode composition uniquement.
- Modification d’[Event.Completed](/javascript/api/office/office.addincommands.event#completed-options-) : ajoute un nouveau paramètre facultatif `options`, qui est un dictionnaire dont la seule valeur valide est `allowEvent`. Cette valeur est utilisée pour annuler l’exécution d’un événement.
- Ajout d’[Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8) : ajoute une nouvelle énumération spécifiant la mise en forme qui s’applique au contenu d’une pièce jointe.
- Ajout d’[Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus?view=outlook-js-1.8) : ajoute une nouvelle énumération qui spécifie si une pièce jointe a été ajoutée à un élément ou supprimée d’un élément.
- Ajout d’[Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor?view=outlook-js-1.8) : ajoute une nouvelle énumération spécifiant les couleurs disponibles à associer à des catégories.
- Ajout d’[Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions?view=outlook-js-1.8) : ajoute une nouvelle énumération d’indicateur binaire spécifiant les autorisations accordées aux délégués.
- Ajout d’[Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype?view=outlook-js-1.8) : ajoute une nouvelle énumération spécifiant le type de lieu d’un rendez-vous.
- Modification d’[Office.EventType](/javascript/api/office/office.eventtype) : ajoute la prise en charge des événements `AttachmentsChanged` et `EnhancedLocationsChanged`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](/outlook/add-ins/quick-start)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
