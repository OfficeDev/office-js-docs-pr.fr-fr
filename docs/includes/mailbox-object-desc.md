Les compléments Outlook utilisent principalement un sous-ensemble de l’API exposée via l’objet [Mailbox](/javascript/api/outlook/office.mailbox). Pour accéder aux objets et aux membres spécifiquement utilisés dans les compléments Outlook, tels que l’objet [Item](/javascript/api/outlook/office.item) , utilisez la propriété de [boîte aux lettres](/javascript/api/office/office.context#office-office-context-mailbox-member) de l’objet **Context** pour accéder à l’objet **Mailbox** , comme indiqué dans la ligne de code suivante.

```js
// Access the Item object.
const item = Office.context.mailbox.item;
```

En outre, les compléments Outlook peuvent utiliser les objets suivants.

- Objet **Office** : pour l’initialisation.

- Objet **Context** : pour l’accès au contenu et aux propriétés de langue d’affichage.

- Objet **RoamingSettings** : pour l’enregistrement des paramètres personnalisés propres au complément Outlook dans la boîte aux lettres de l’utilisateur dans laquelle le complément est installé.

Pour plus d’informations sur l’utilisation de JavaScript dans les compléments Outlook, reportez-vous à la rubrique [Compléments Outlook](../outlook/outlook-add-ins-overview.md).
