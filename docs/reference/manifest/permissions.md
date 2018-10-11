# <a name="permissions-element"></a>Élément Permissions

Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.

**Type de complément :** contenu, volet Office, messagerie

## <a name="syntax"></a>Syntaxe

Pour les compléments du volet de tâches et de contenu :

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Pour les compléments de messagerie

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Remarques

Pour plus de détails, consultez l’article relatif à la [demande d’autorisations pour utiliser des API dans des compléments de contenu et de volet Office](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) et celui décrivant les [autorisations de complément Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).
