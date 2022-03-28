Outlook de messagerie utilisent principalement les API exposées via [l’objet Mailbox](/javascript/api/outlook/office.mailbox). Pour accéder aux objets et aux membres destinés spécifiquement à une utilisation dans les compléments Outlook, tels que l’objet [Item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item), utilisez la propriété [mailbox](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox) de l’objet **Context** pour accéder à l’objet **Mailbox**, comme illustré dans la ligne de code suivante.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

En outre, Outlook compléments peuvent utiliser les objets suivants.

-  Objet **Office** : pour l’initialisation.

-  Objet **Context** : pour l’accès au contenu et aux propriétés de langue d’affichage.

-  Objet **RoamingSettings** : pour l’enregistrement des paramètres personnalisés propres au complément Outlook dans la boîte aux lettres de l’utilisateur dans laquelle le complément est installé.

Pour plus d’informations sur l’utilisation Outlook’API JavaScript, voir [Outlook de l’api](../outlook/outlook-add-ins-overview.md) JavaScript.