Les compléments Outlook utilisent principalement les API exposées par le biais de l’objet [Mailbox](/javascript/api/outlook/Office.mailbox) . Pour accéder aux objets et aux membres destinés spécifiquement à une utilisation dans les compléments Outlook, tels que l’objet [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md), utilisez la propriété [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) de l’objet **Context** pour accéder à l’objet **Mailbox**, comme illustré dans la ligne de code suivante.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

De plus, les compléments Outlook peuvent utiliser les objets suivants :

-  Objet **Office** : pour l’initialisation.

-  Objet **Context** : pour l’accès au contenu et aux propriétés de langue d’affichage.

-  Objet **RoamingSettings** : pour l’enregistrement des paramètres personnalisés propres au complément Outlook dans la boîte aux lettres de l’utilisateur dans laquelle le complément est installé.

Pour plus d’informations sur l’utilisation de l’API JavaScript Outlook, consultez la rubrique [compléments Outlook](../outlook/outlook-add-ins-overview.md).