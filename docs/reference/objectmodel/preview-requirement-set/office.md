 

# <a name="office"></a>Office

L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale requise de la boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="members-and-methods"></a>Membres et méthodes

| Membre | Type |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Membre |
| [CoercionType](#coerciontype-string) | Membre |
| [EventType](#eventtype-string) | Membre |
| [SourceProperty](#sourceproperty-string) | Membre |

### <a name="namespaces"></a>Espaces de noms

[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.

### <a name="members"></a>Membres

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Spécifie le résultat d’un appel asynchrone.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Succeeded`| Chaîne|L’appel a réussi.|
|`Failed`| Chaîne|L’appel n’a pas réussi.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale requise de la boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

---

####  <a name="coerciontype-string"></a>CoercionType :String

Indique comment forcer le type des données retournées ou définies par la méthode appelée.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Html`| Chaîne|Demande que les données soient renvoyées au format HTML.|
|`Text`| Chaîne|Demande que les données soient renvoyées au format texte.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale requise de la boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

---

####  <a name="eventtype-string"></a>EventType :String

Spécifie l’événement associé à un gestionnaire d’événements.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="properties"></a>Propriétés :

| Nom | Type | Description | Ensemble minimal de conditions requises |
|---|---|---|---|
|`AppointmentTimeChanged`| Chaîne | La date ou l’heure du rendez-vous sélectionné ou de la série a changé. | 1.7 |
|`ItemChanged`| Chaîne | L’élément sélectionné a changé. | 1.5 |
|`OfficeThemeChanged`| Chaîne | L’élément sélectionné a changé. | Aperçu |
|`RecipientsChanged`| Chaîne | La liste de destinataires de l’élément sélectionné ou l’emplacement du rendez-vous a changé. | 1.7 |
|`RecurrenceChanged`| Chaîne | La périodicité de la série sélectionnée a changé. | 1.7 |

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale requise de la boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

Spécifie la source des données renvoyées par la méthode appelée.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Body`| Chaîne|La source de données est dans le corps d’un message.|
|`Subject`| Chaîne|La source de données est dans l’objet d’un message.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale requise de la boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|