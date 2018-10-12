 

# <a name="office"></a>Office

L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

### <a name="namespaces"></a>Espaces de noms

[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.

### <a name="members"></a>Membres

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus : Chaîne

Spécifie le résultat d’un appel asynchrone.

##### <a name="type"></a>Type :

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Succeeded`| String|L’appel a réussi.|
|`Failed`| String|L’appel n’a pas réussi.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|
####  <a name="coerciontype-string"></a>CoercionType : Chaîne

Indique comment forcer le type des données retournées ou définies par la méthode appelée.

##### <a name="type"></a>Type :

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Html`| String|Demande que les données soient renvoyées au format HTML.|
|`Text`| String|Demande que les données soient renvoyées au format texte.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|
####  <a name="sourceproperty-string"></a>SourceProperty : Chaîne

Spécifie la source des données renvoyées par la méthode appelée.

##### <a name="type"></a>Type :

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Body`| String|La source de données est dans le corps d’un message.|
|`Subject`| String|La source de données est dans l’objet d’un message.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|