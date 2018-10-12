
# <a name="userprofile"></a>userProfile

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale requise de la boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="members-and-methods"></a>Membres et méthodes

| Membre | Type |
|--------|------|
| [accountType](#accounttype-string) | Membre |
| [displayName](#displayname-string) | Membre |
| [emailAddress](#emailaddress-string) | Membre |
| [timeZone](#timezone-string) | Membre |

### <a name="members"></a>Membres

####  <a name="accounttype-string"></a>accountType :String

> [!NOTE]
> Ce membre est actuellement uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (build 16.9.1212 ou version ultérieure).

Obtient le type de compte de l’utilisateur associé à la boîte aux lettres. Les valeurs possibles sont répertoriées dans le tableau suivant.

| Valeur | Description |
|-------|-------------|
| `enterprise` | La boîte aux lettres est sur un serveur Exchange local. |
| `gmail` | La boîte aux lettres est associée à un compte Gmail. |
| `office365` | La boîte aux lettres est associée avec un compte Office 365 professionnel ou scolaire. |
| `outlookCom` | La boîte aux lettres est associée à un compte Outlook.com personnel. |

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale requise de la boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a>displayName :String

Obtient le nom d’affichage de l’utilisateur.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale requise de la boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a>emailAddress :String

Obtient l’adresse e-mail SMTP de l’utilisateur.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale requise de la boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a>timeZone :String

Obtient le fuseau horaire par défaut de l’utilisateur.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale requise de la boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```