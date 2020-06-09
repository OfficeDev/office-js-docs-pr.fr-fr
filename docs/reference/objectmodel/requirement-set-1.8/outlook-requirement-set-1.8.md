---
title: Ensemble de conditions requises de l’API du complément Outlook 1.8
description: Ensemble de conditions requises 1,8 pour l’API de complément Outlook
ms.date: 05/20/2020
localization_priority: Normal
ms.openlocfilehash: b0ccf50d93c94787d9e8349324c7cc3705d9223d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612170"
---
# <a name="outlook-add-in-api-requirement-set-18"></a>Ensemble de conditions requises de l’API du complément Outlook 1.8

Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.

## <a name="whats-new-in-18"></a>Nouveautés de la version 1.8

L’ensemble de conditions requises de la version 1.8 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md). Les fonctionnalités suivantes ont été ajoutées.

- Ajout de nouvelles fonctionnalités d’API pour les pièces jointes, de catégories, d’accès délégué, d’emplacement amélioré, d’en-têtes Internet et de blocage d’envoi.
- Ajout d’un paramètre `options` facultatif à Event.completed.
- Ajout de la prise en charge de `AttachmentsChanged` et des `EnhancedLocationsChanged` événements.

### <a name="change-log"></a>Journal des modifications

- Ajout d’[AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8) : ajoute un nouvel objet représentant le contenu d’une pièce jointe.
- Ajout de [AttachmentDetailsCompose](/javascript/api/outlook/office.attachmentdetailscompose?view=outlook-js-1.8): ajoute un nouvel objet qui représente les détails d’une pièce jointe en mode composition.
- Ajout de [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8) : ajoute un nouvel objet représentant les catégories d’un élément.
- Ajout de [CategoryDetails](/javascript/api/outlook/office.categorydetails?view=outlook-js-1.8) : ajoute un nouvel objet représentant les détails d’une catégorie (son nom et la couleur associée).
- Ajout d’[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) : ajoute un nouvel objet représentant l’ensemble des lieux pour un rendez-vous.
- Ajout d’[InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8) : ajout d’un nouvel objet représentant les en-têtes Internet d’un élément de message. Mode composition uniquement.
- Ajout de [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) : ajoute un nouvel objet représentant un lieu. En lecture seule.
- Ajout de [LocationIdentifier](/javascript/api/outlook/office.locationidentifier?view=outlook-js-1.8) : ajoute un nouvel objet représentant l’ID d’un lieu.
- Ajout de [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8) : ajoute un nouvel objet représentant la liste principale des catégories d’une boîte aux lettres.
- Ajout de [SharedProperties](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) : ajoute un nouvel objet représentant les propriétés d’un élément de rendez-vous ou de message dans un dossier, un calendrier ou une boîte aux lettres partagé(e).
- Ajout d’un [élément de manifeste SupportsSharedFolders](../../manifest/supportssharedfolders.md) : ajoute un élément enfant à l’élément de manifeste [DesktopFormFactor](../../manifest/desktopformfactor.md). Définit si le complément est disponible dans les scénarios de délégué.
- Ajout d’[Office.context.mailbox.masterCategories](office.context.mailbox.md#properties) : ajoute une nouvelle propriété représentant la liste principale des catégories d’une boîte aux lettres.
- Ajout d’[Office.context.mailbox.item.categories](office.context.mailbox.item.md#properties) : ajoute une nouvelle propriété représentant l’ensemble des catégories d’une boîte aux lettres.
- Ajout d’[Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#methods) : ajoute une nouvelle méthode qui vous permet de joindre un fichier à un message ou à un rendez-vous. Ce fichier est représenté par une chaîne encodée en base 64.
- Ajout d’[Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#properties) : ajoute une nouvelle propriété représentant l’ensemble des lieux pour un rendez-vous.
- Ajout d'[Office.context.mailbox.item.getAllInternetHeadersAsync](office.context.mailbox.item.md#methods) : ajoute une nouvelle méthode récupérant tous les en-têtes Internet pour un élément de message. Mode Lecture uniquement.
- Ajout d’[Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#methods) : ajoute une nouvelle méthode pour obtenir le contenu d’une pièce jointe spécifique.
- Ajout d’[Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#methods) : ajoute une nouvelle méthode récupérant les pièces jointes à un élément en mode composition.
- Ajout d’[Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#methods) : ajoute une nouvelle méthode obtenant l’ID d’un rendez-vous ou d’un élément de message enregistré.
- Ajout d’[Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#methods) : ajout d’une nouvelle méthode obtenant un objet représentant les sharedProperties d’un rendez-vous ou d’un élément de message.
- Ajout d’[Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#properties) : ajoute une nouvelle propriété représentant les en-têtes Internet personnalisés d’élément de message. Mode composition uniquement.
- Modification d’[Event.Completed](/javascript/api/office/office.addincommands.event#completed-options-) : ajoute un nouveau paramètre facultatif `options`, qui est un dictionnaire dont la seule valeur valide est `allowEvent`. Cette valeur est utilisée pour annuler l’exécution d’un événement.
- Ajout d’[Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8) : ajoute une nouvelle énumération spécifiant la mise en forme qui s’applique au contenu d’une pièce jointe.
- Ajout d’[Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus?view=outlook-js-1.8) : ajoute une nouvelle énumération qui spécifie si une pièce jointe a été ajoutée à un élément ou supprimée d’un élément.
- Ajout d’[Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor?view=outlook-js-1.8) : ajoute une nouvelle énumération spécifiant les couleurs disponibles à associer à des catégories.
- Ajout d’[Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions?view=outlook-js-1.8) : ajoute une nouvelle énumération d’indicateur binaire spécifiant les autorisations accordées aux délégués.
- Ajout d’[Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype?view=outlook-js-1.8) : ajoute une nouvelle énumération spécifiant le type de lieu d’un rendez-vous.
- Modification d’[Office.EventType](/javascript/api/office/office.eventtype) : ajoute la prise en charge des événements `AttachmentsChanged` et `EnhancedLocationsChanged`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
