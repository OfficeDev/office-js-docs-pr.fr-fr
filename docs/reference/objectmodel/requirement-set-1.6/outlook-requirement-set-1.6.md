# <a name="outlook-add-in-api-requirement-set-16"></a>Ensemble de conditions requises de l’API de complément Outlook 1.6

Le sous-ensemble de l’API de complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises précédent.

## <a name="whats-new-in-16"></a>Nouveautés de la version 1.6

L’ensemble de conditions requises 1.6 inclut toutes les fonctionnalités de l’[ensemble de conditions requises 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md). Les fonctionnalités suivantes ont été ajoutées.

- Ajout de nouvelles API pour les compléments contextuelles pour obtenir de l’entité ou la correspondance RegEx que l’utilisateur a sélectionné pour activer le complément.
- Ajout d’une nouvelle API pour ouvrir un formulaire de nouveau message.
- Ajout de la possibilité pour le complément de déterminer le type de compte de boîte aux lettres de l’utilisateur.

### <a name="change-log"></a>Journal des modifications

- Ajout de [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities) : ajout d’une fonction qui obtient les entités figurant dans une correspondance en surbrillance sélectionnée par un utilisateur. Les correspondances en surbrillance s’appliquent aux compléments contextuels.
- Ajout de [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object) : ajout d’une fonction qui renvoie les valeurs de chaîne dans une correspondance en surbrillance qui correspondent aux expressions régulières définies dans le fichier XML de manifeste. Les correspondances en surbrillance s’appliquent aux compléments contextuels.
- Ajout de [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters) : ajoute une nouvelle fonction qui ouvre un formulaire de nouveau message.
- Ajout de [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string) : ajoute un nouveau membre au profil d’utilisateur qui indique le type de compte de l’utilisateur.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)