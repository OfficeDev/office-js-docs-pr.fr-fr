# <a name="outlook-add-in-api-preview-requirement-set"></a>Ensemble de conditions requises de l’API du complément Outlook (aperçu)

Le sous-ensemble de l’API de complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Cette documentation est pour un **Aperçu** [exigence](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). Cet ensemble de conditions requises n’est pas totalement implémenté encore et clients ne signaleront pas correctement prise en charge pour celui-ci. Vous ne devez pas spécifier cette exigence définie dans le manifeste de votre complément. Méthodes et propriétés qui sont introduites dans cet ensemble de conditions requises doivent être individuellement tester disponibilité avant de les utiliser.

L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Fonctionnalités (aperçu) :

Les fonctionnalités suivantes sont disponibles en aperçu.

- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - Ajout d’un nouvel objet qui représente les propriétés d’un rendez-vous ou d’un élément de message dans un dossier partagé, un calendrier ou une boîte aux lettres.
- [Event.Completed](/javascript/api/office/office.addincommands.event#completed-options-) : nouveau paramètre facultatif `options`, qui est un dictionnaire ayant comme seule valeur valide `allowEvent`. Cette valeur est utilisée pour annuler l’exécution d’un événement.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - Ajout d’une nouvelle méthode qui joint un fichier issu de l'encodage base64 à un message ou à un rendez-vous.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) - Ajout d’une fonction qui renvoie les données d’initialisation transmises quand le complément est [activé par un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - Ajout d’une nouvelle méthode qui obtient un objet qui représente la propriété sharedProperties d’un rendez-vous ou d’un élément de message.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) - Ajout d’un accès à `getAccessTokenAsync`, qui permet aux compléments d’[obtenir un jeton d’accès](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour l’API Microsoft Graph.
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - Ajout d’une nouvelle valeur de bit indicateur qui spécifie les autorisations accordées au délégué.
- [Office.EventType](/javascript/api/office/office.eventtype) - Modifié pour prendre en charge l’événement OfficeThemeChanged grâce à l’ajout de l’entrée `OfficeThemeChanged`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)