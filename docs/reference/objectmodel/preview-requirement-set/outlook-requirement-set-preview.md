---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: b46fada2fa69f3526c929a0289341f7dab5b58b8
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128474"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Ensemble de conditions requises de l’API du complément Outlook (aperçu)

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Cette documentation a trait à un [ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) en **préversion**. Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions. Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément. La disponibilité des méthodes et des propriétés présentées dans cet ensemble de conditions doit être testée avant de les utiliser. Vous devrez également participer au [programme Office Insider](https://products.office.com/office-insider).

L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Fonctionnalités (aperçu) :

Les fonctionnalités suivantes sont disponibles en aperçu.

### <a name="attachments"></a>Pièces jointes

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

Ajout d’un nouvel objet représentant le contenu d’une pièce jointe.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

Ajout d’une nouvelle méthode qui vous permet de joindre un fichier représenté par une chaîne encodée en base 64 à un message ou à un rendez-vous.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

Ajout d’une nouvelle méthode pour obtenir le contenu d’une pièce jointe spécifique.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

Ajout d’une nouvelle méthode qui obtient les pièces jointes d’un élément en mode composition.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

Ajout d’une nouvelle énumération qui spécifie la mise en forme qui s’applique au contenu d’une pièce jointe.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

Ajout d’une nouvelle énumération qui spécifie si une pièce jointe a été ajoutée à un élément ou supprimée d’un élément.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.AttachmentsChanged](/javascript/api/office/office.eventtype)

Ajout de l’événement `AttachmentsChanged` à `Item`.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

---

### <a name="block-on-send"></a>Blocage lors de l’envoi

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[Event.completed](/javascript/api/office/office.addincommands.event#completed-options-)

Ajout d’un nouveau paramètre facultatif `options`, qui est un dictionnaire ayant comme seule valeur valide `allowEvent`. Cette valeur est utilisée pour annuler l’exécution d’un événement.

**Disponible dans** : Outlook sur le web (classique)

---

### <a name="categories"></a>Catégories

Dans Outlook, un utilisateur peut regrouper des messages et des rendez-vous à l’aide d’une catégorie pour leur appliquer un code de couleur. L’utilisateur définit les catégories dans une liste sur sa boîte aux lettres principale. Ils peuvent ensuite appliquer une ou plusieurs catégories à un élément.

> [!NOTE]
> Cette fonctionnalité n’est pas prise en charge dans Outlook sur iOS ou Android.

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[Categories](/javascript/api/outlook/office.categories)

Ajout d’un nouvel objet représentant des catégories d’un élément.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[Détailscatégorie](/javascript/api/outlook/office.categorydetails)

Ajout un nouvel objet qui représente les détails d’une catégorie (son nom et la couleur associée).

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[Catégoriesmaître](/javascript/api/outlook/office.mastercategories)

Ajout d’ un nouvel objet qui représente la liste Catégories maître sur une boîte aux lettres.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[Office.context.mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)

Ajout d’ un nouveau propriétaire qui représente la liste Catégories maître sur une boîte aux lettres.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[Office.context.mailbox.item.categories](/javascript/api/outlook/office.item#categories)

Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un élément.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor)

Ajouté un nouvel enum qui spécifie les couleurs disponibles à associer à des catégories.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

---

### <a name="delegate-access"></a>Accès délégué

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[SharedProperties](/javascript/api/outlook/office.sharedproperties)

Ajout d’un nouvel objet qui représente les propriétés d’un élément rendez-vous ou message dans un dossier, un calendrier ou une boîte aux lettres partagés.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#getitemidasyncoptions-callback)

Ajout d’une nouvelle méthode qui obtient l’ID d’un rendez-vous ou d’un élément de message enregistré.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

Ajout d’une nouvelle méthode qui obtient un objet qui représente les sharedProperties d’un élément rendez-vous ou message.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

Ajout d’une nouvelle énumération d’indicateur binaire qui spécifie les autorisations accordées aux délégués.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[Élément de manifeste SupportsSharedFolders](../../manifest/supportssharedfolders.md)

Ajout d’un élément enfant à l’élément de manifeste [DesktopFormFactor](../../manifest/desktopformfactor.md). Définit si le complément est disponible dans les scénarios de délégué.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

---

### <a name="enhanced-location"></a>Emplacement amélioré

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

Ajout d’un nouvel objet représentant l’ensemble des emplacements sur un rendez-vous.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[LocationDetails](/javascript/api/outlook/office.locationdetails)

Ajout d’un nouvel objet représentant un emplacement. En lecture seule.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[LocationIdentifier](/javascript/api/outlook/office.locationidentifier)

Ajout d’un nouvel objet représentant l’ID d’un emplacement.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

Ajout d’une nouvelle propriété représentant l’ensemble des emplacements sur un rendez-vous.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)

Ajout d’une nouvelle énumération qui spécifie le type d’emplacement d’un rendez-vous.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.EnhancedLocationsChanged](/javascript/api/office/office.eventtype)

Ajout de l’événement `EnhancedLocationsChanged` à `Item`.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

---

### <a name="integration-with-actionable-messages"></a>Intégration avec les messages actionnables

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (classique)

---

### <a name="internet-headers"></a>En-têtes Internet

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[InternetHeaders](/javascript/api/outlook/office.internetheaders)

Ajout d’un nouvel objet représentant les en-têtes Internet d’un élément de message.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheaders)

Ajout d’une nouvelle propriété représentant les en-têtes Internet d’un élément de message.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

---

### <a name="office-theme"></a>Thème Office

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[Office.context.mailbox.officeTheme](/javascript/api/office/office.officetheme)

Ajout de la possibilité d’obtenir un thème Office.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

---

### <a name="sso"></a>Authentification unique

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[Office.context.auth.getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

Ajout d’un accès à `getAccessTokenAsync`, qui permet aux compléments d’[obtenir un jeton d’accès](/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour l’API Microsoft Graph.

**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365), Outlook sur le web (nouveau), Outlook sur le web (classique)

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](/outlook/add-ins/quick-start)
