|**Nom canonique au niveau de l’autorisation</br>**|**Nom du manifeste XML**|**Nom du manifeste Teams**|**Description récapitulative**|
|:-----|:-----|:-----|:-----|
|**Restreint**|Restreint|MailboxItem.Restricted.User|Autorise l’utilisation d’entités, mais pas d’expressions régulières. |
|**lire l’élément**|ReadItem|MailboxItem.Read.User|En plus de ce qui est autorisé dans **restreint**, il autorise :<ul><li>expressions régulières</li><li>l’accès en lecture de l’API du complément Outlook</li><li>l’obtention des propriétés de l’élément et du jeton de rappel</li></ul> |
|**élément en lecture/écriture**|ReadWriteItem|MailboxItem.ReadWrite.User|En plus de ce qui est autorisé dans **l’élément de lecture**, il permet :<ul><li>l’accès total à l’API du complément Outlook, à l’exception de `makeEwsRequestAsync`</li><li>la définition des propriétés de l’élément</li></ul> |
|**boîte aux lettres en lecture/écriture**|ReadWriteMailbox|Mailbox.ReadWrite.User|En plus de ce qui est autorisé dans **l’élément de lecture/écriture**, il permet :<ul><li>la création, la lecture, l’écriture d’éléments et de dossiers</li><li>l’envoi d’éléments</li><li>l’appel de [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)</li></ul> |

Les autorisations sont déclarées dans le manifeste. Le balisage varie en fonction du type de manifeste.

- **Manifeste XML** : utilisez l’élément **\<Permissions\>** .
- **Manifeste Teams (préversion)** : utilisez la propriété « name » d’un objet dans le tableau « authorization.permissions.resourceSpecific ».

> [!NOTE]
>
> - Une autorisation supplémentaire est nécessaire pour les compléments qui utilisent la fonctionnalité d’ajout à l’envoi. Avec le manifeste XML, vous spécifiez l’autorisation dans l’élément [ExtendedPermissions](/javascript/api/manifest/extendedpermissions) . Pour plus d’informations, consultez [Implémenter l’ajout à l’envoi dans votre complément Outlook](../outlook/append-on-send.md). Avec le manifeste Teams (préversion), vous spécifiez cette autorisation avec le nom **Mailbox.AppendOnSend.User** dans un objet supplémentaire dans le tableau « authorization.permissions.resourceSpecific ».
> - Une autorisation supplémentaire est nécessaire pour les compléments qui utilisent des dossiers partagés. Avec le manifeste XML, vous spécifiez l’autorisation en définissant l’élément [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) sur `true`. Pour plus d’informations, consultez [Activer les dossiers partagés et les scénarios de boîte aux lettres partagées dans un complément Outlook](../outlook/delegate-access.md). Avec le manifeste Teams (préversion), vous spécifiez cette autorisation avec le nom **Mailbox.SharedFolder** dans un objet supplémentaire dans le tableau « authorization.permissions.resourceSpecific ».
