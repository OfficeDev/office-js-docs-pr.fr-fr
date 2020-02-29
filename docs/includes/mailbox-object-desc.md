<span data-ttu-id="6db4e-101">Les compléments Outlook utilisent principalement les API exposées par le biais de l’objet [Mailbox](/javascript/api/outlook/Office.mailbox) .</span><span class="sxs-lookup"><span data-stu-id="6db4e-101">Outlook add-ins primarily use the APIs exposed through the [Mailbox](/javascript/api/outlook/Office.mailbox) object.</span></span> <span data-ttu-id="6db4e-102">Pour accéder aux objets et aux membres destinés spécifiquement à une utilisation dans les compléments Outlook, tels que l’objet [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md), utilisez la propriété [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) de l’objet **Context** pour accéder à l’objet **Mailbox**, comme illustré dans la ligne de code suivante.</span><span class="sxs-lookup"><span data-stu-id="6db4e-102">To access the objects and members specifically for use in Outlook add-ins, such as the [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) object, you use the [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) property of the **Context** object to access the **Mailbox** object, as shown in the following line of code.</span></span>

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

<span data-ttu-id="6db4e-103">De plus, les compléments Outlook peuvent utiliser les objets suivants :</span><span class="sxs-lookup"><span data-stu-id="6db4e-103">Additionally, Outlook add-ins can use the following objects:</span></span>

-  <span data-ttu-id="6db4e-104">Objet **Office** : pour l’initialisation.</span><span class="sxs-lookup"><span data-stu-id="6db4e-104">**Office** object: for initialization.</span></span>

-  <span data-ttu-id="6db4e-105">Objet **Context** : pour l’accès au contenu et aux propriétés de langue d’affichage.</span><span class="sxs-lookup"><span data-stu-id="6db4e-105">**Context** object: for access to content and display language properties.</span></span>

-  <span data-ttu-id="6db4e-106">Objet **RoamingSettings** : pour l’enregistrement des paramètres personnalisés propres au complément Outlook dans la boîte aux lettres de l’utilisateur dans laquelle le complément est installé.</span><span class="sxs-lookup"><span data-stu-id="6db4e-106">**RoamingSettings** object: for saving Outlook add-in-specific custom settings to the user's mailbox where the add-in is installed.</span></span>

<span data-ttu-id="6db4e-107">Pour plus d’informations sur l’utilisation de l’API JavaScript Outlook, consultez la rubrique [compléments Outlook](../outlook/outlook-add-ins-overview.md).</span><span class="sxs-lookup"><span data-stu-id="6db4e-107">For information about using the Outlook JavaScript API, see [Outlook add-ins](../outlook/outlook-add-ins-overview.md).</span></span>