---
title: Faire correspondre les chaînes en tant qu’entités connues dans un complément Outlook
description: À l’Office’API JavaScript, vous pouvez obtenir des chaînes qui correspondent à des entités connues spécifiques pour un traitement ultérieur.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 8d4b78259b771d29244641d9e3ca867018b763ef
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348495"
---
# <a name="match-strings-in-an-outlook-item-as-well-known-entities"></a>Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues

Avant d’envoyer un élément de message ou de demande de réunion, Exchange Server analyse le contenu de l’élément, identifie et marque certaines chaînes dans l’objet et le corps similaires aux entités connues d’Exchange (par exemple, adresses e-mail, numéros de téléphone, URL). Les demandes de réunion et les messages sont envoyés par Exchange Server dans une boîte de réception Outlook avec les entités connues marquées. 

À l’Office’API JavaScript, vous pouvez obtenir ces chaînes qui correspondent à des entités connues spécifiques pour un traitement ultérieur. Vous pouvez également spécifier une entité connue dans une règle du manifeste du complément pour qu’Outlook puisse activer votre complément quand l’utilisateur affiche un élément contenant des correspondances pour cette entité. Vous pouvez extraire et effectuer une action sur des correspondances pour cette entité. 

Pouvoir identifier ou extraire de telles instances à partir d’un message ou d’un rendez-vous sélectionné s’avère très pratique. Par exemple, vous pouvez créer un service de recherche sur annuaire inversé comme complément Outlook. Le complément peut extraire des chaînes dans l’objet ou le corps de l’élément semblable à un numéro de téléphone, effectuer une recherche sur annuaire inversé et afficher le propriétaire inscrit de chaque numéro de téléphone.

Cette rubrique présente ces entités connues, montre des exemples de règles d’activation en fonction de ces entités et explique comment extraire des correspondances d’entités indépendamment de l’utilisation d’entités dans les règles d’activation.


## <a name="support-for-well-known-entities"></a>Prise en charge des entités connues

Exchange Server marque les entités connues dans un élément de message ou de demande de réunion après que l’élément a été envoyé par l’expéditeur et avant qu’il soit remis au destinataire. Ainsi, seuls les éléments ayant transité via Exchange sont marqués, et Outlook peut activer des compléments en fonction de ces marquages quand l’utilisateur affiche ces éléments. En revanche, quand l’utilisateur compose ou affiche un élément du dossier Éléments envoyés, Outlook ne peut pas activer les compléments en fonction des entités connues car l’élément n’a pas transité via Exchange. 

De même, vous ne pouvez pas extraire les entités connues dans les éléments en cours de composition ou situés dans le dossier Éléments envoyés, car ces éléments n’ont pas transité via Exchange et ne sont pas marqués. Pour plus d’informations sur les types d’éléments qui prennent en charge l’activation, voir [Règles d’activation pour les compléments Outlook](activation-rules.md).

Le tableau suivant répertorie les entités qu’Exchange Server et Outlook prennent en charge et reconnaissent (d’où le nom « entités connues »), et le type d’objet d’une instance de chaque entité. La reconnaissance du langage naturel d’une chaîne en tant que l’une de ces entités est fondée sur un modèle d’apprentissage qui a été testé sur une grande quantité de données. Par conséquent, la reconnaissance n’est pas déterministe. Pour plus d’informations sur les conditions de reconnaissance, voir [Conseils d’utilisation des entités connues](#tips-for-using-well-known-entities).

**Tableau 1. Entités prises en charge et leurs types**

|Type d’entité|Conditions de reconnaissance|Type d’objet|
|:-----|:-----|:-----|
|**Adresse**|Adresses aux États-Unis ; par exemple : 1234 Main Street, Redmond, WA 07722. En général, pour qu’une adresse soit reconnue, elle doit suivre la structure d’une adresse postale des États-Unis, où la plupart des éléments sont présents, à savoir numéro de rue, nom de rue, ville, État et code postal. L’adresse peut être spécifiée sur une ou plusieurs lignes.|Objet JavaScript **String**|
|**Contact**|Une référence aux informations d’une personne telles que reconnues en langage naturel. La reconnaissance d’un contact varie selon le contexte. Par exemple, une signature à la fin d’un message ou le nom d’une personne apparaissant à proximité des informations suivantes : un numéro de téléphone, une adresse, une adresse électronique et une URL.|Objet [Contact](/javascript/api/outlook/office.contact)|
|**EmailAddress**|Adresses électroniques SMTP.|Objet `String` JavaScript|
|**MeetingSuggestion**|Une référence à un événement ou une réunion. Par exemple, Exchange 2013 reconnaîtrait le texte suivant comme une suggestion de réunion :  _On se voit demain pour déjeuner ?_|Objet [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|
|**PhoneNumber**|Numéros de téléphone des États-Unis ; par exemple :  _(235) 555-0110_|Objet [PhoneNumber](/javascript/api/outlook/office.phonenumber)|
|**TaskSuggestion**|Phrases appelant une action. Par exemple :  _Veuillez mettre à jour la feuille de calcul._|Objet [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)|
|**Url**|Adresse web qui spécifie explicitement l’identificateur et l’emplacement réseau d’une ressource web. Exchange Server ne nécessite pas le protocole d’accès dans l’adresse web et ne reconnaît pas les URL incorporées dans le texte du lien en tant qu’instances de `Url` l’entité. Exchange Server peuvent correspondre aux exemples suivants : `www.youtube.com/user/officevideos``https://www.youtube.com/user/officevideos` |Objet `String` JavaScript|

<br/>

La figure suivante décrit comment Exchange Server et Outlook prennent en charge les entités connues pour les compléments et indique ce que les compléments peuvent faire avec ces entités connues. Reportez-vous à [Récupération d’entités dans votre complément](#retrieving-entities-in-your-add-in) et [Activation d’un complément basé sur l’existence d’une entité](#activating-an-add-in-based-on-the-existence-of-an-entity) pour plus de détails sur l’utilisation de ces entités.

**Prise en charge des entités connues par Exchange Server, Outlook et les compléments**

![Prise en charge et utilisation d’entités connues dans l’application de messagerie.](../images/well-known-entities-info.png)


## <a name="permissions-to-extract-entities"></a>Autorisations d’extraction d’entités

Pour extraire les entités de votre code JavaScript ou pour activer votre complément à partir de l’existence de certaines entités connues, assurez-vous que vous avez demandé les autorisations appropriées dans le manifeste du complément.

La spécification de l’autorisation restreinte par défaut permet à votre add-in d’extraire `Address` le `MeetingSuggestion` , ou `TaskSuggestion` l’entité. Pour extraire les autres entités, spécifiez les autorisations de lecture d’élément, de lecture/écriture d’élément ou de lecture/écriture de boîte aux lettres. Pour ce faire dans le manifeste, utilisez l’élément [Permissions](../reference/manifest/permissions.md) et spécifiez l’autorisation appropriée &mdash; **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox** comme dans l’exemple &mdash; suivant.

```xml
<Permissions>ReadItem</Permissions>
```


## <a name="retrieving-entities-in-your-add-in"></a>Récupération d’entités dans votre complément

Tant que l’objet ou le corps de l’élément consulté par l’utilisateur contient des chaînes qu’Exchange et Outlook peuvent reconnaître comme des entités connues, ces instances sont disponibles pour les compléments, et ce même si un complément n’est pas activé en fonction des entités connues. Avec l’autorisation appropriée, vous pouvez utiliser la ou la méthode pour récupérer les entités connues présentes dans le message ou le `getEntities` `getEntitiesByType` rendez-vous actuel.

La méthode renvoie un tableau d’objets Entities qui contient toutes les entités `getEntities` connues dans [](/javascript/api/outlook/office.entities) l’élément.

Si vous êtes intéressé par un type particulier d’entités, utilisez la méthode qui renvoie uniquement un tableau des `getEntitiesByType` entités de votre choix. L’énumération [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) représente tous les types d’entités connues que vous pouvez extraire.

Après avoir appelé, vous pouvez ensuite utiliser la propriété correspondante de l’objet pour obtenir un tableau `getEntities` `Entities` d’instances d’un type d’entité. Selon le type d’entité, les instances dans le tableau peuvent être seulement des chaînes, ou peuvent être mappés avec des objets spécifiques. 

Comme dans l’exemple illustré dans la figure précédente, pour obtenir des adresses dans l’élément, accédez au tableau renvoyé par `getEntities().addresses[]`. La `Entities.addresses` propriété renvoie un tableau de chaînes qui Outlook comme adresses postales. De même, la propriété renvoie un tableau d’objets que Outlook `Entities.contacts` `Contact` reconnaît en tant qu’informations de contact. Le tableau 1 répertorie le type d’objet d’une instance de chaque entité prise en charge.

L’exemple suivant illustre comment récupérer des adresses trouvées dans un message.

```js
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities && null != entities.addresses && undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## <a name="activating-an-add-in-based-on-the-existence-of-an-entity"></a>Activation d’un complément basé sur l’existence d’une entité

Une autre façon d’utiliser des entités connues consiste à faire en sorte qu’Outlook active votre complément selon l’existence de types d’entités dans l’objet ou le corps de l’élément actuellement affiché. Vous pouvez le faire en spécifiant une `ItemHasKnownEntity` règle dans le manifeste du add-in. Le type simple [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) représente les différents types d’entités connues pris en charge par les `ItemHasKnownEntity` règles. Une fois votre complément activé, vous pouvez également récupérer les instances de ces entités pour répondre à vos besoins, comme le décrit la section précédente [Récupération d’entités dans votre complément](#retrieving-entities-in-your-add-in).

Vous pouvez éventuellement appliquer une expression régulière dans une règle, de manière à filtrer davantage les instances d’une entité et demander à Outlook d’activer un add-in uniquement sur un sous-ensemble des instances de `ItemHasKnownEntity` l’entité. Par exemple, vous pouvez spécifier un filtre pour l’entité d’adresse dans un message qui contient un code postal de l’État de Washington commençant par « 98 ». Pour appliquer un filtre sur les instances d’entité, utilisez les attributs dans l’élément du `RegExFilter` `FilterName` type `Rule` [ItemHasKnownEntity.](../reference/manifest/rule.md#itemhasknownentity-rule)

Comme avec d’autres règles d’activation, vous pouvez spécifier plusieurs règles afin de former une collection de règles pour votre complément. L’exemple suivant applique une opération « AND » sur 2 règles : une `ItemIs` règle et une `ItemHasKnownEntity` règle. Cette collection de règles active le complément lorsque l’élément en cours est un message et qu’Outlook reconnaît une adresse dans l’objet ou le corps de cet élément.

```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<br/>

L’exemple suivant utilise l’élément actuel pour définir une variable sur `getEntitiesByType` les résultats de la collection de règles `addresses` précédente.

```js
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

<br/>

L’exemple de règle suivant active le add-in chaque fois qu’il existe une URL dans l’objet ou le corps de l’élément actif, et que l’URL contient la chaîne « youtube », quelle que soit la valeur de la `ItemHasKnownEntity` chaîne.

```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

<br/>

L’exemple suivant utilise l’élément actuel pour définir une variable afin d’obtenir un tableau de résultats qui correspondent à `getFilteredEntitiesByName(name)` l’expression régulière de la règle `videos` `ItemHasKnownEntity` précédente.

```js
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## <a name="tips-for-using-well-known-entities"></a>Conseils d’utilisation des entités connues

Si vous utilisez des entités connues dans votre complément, vous devez connaître certaines informations et limites. L’exemple suivant s’applique tant que votre application est activée lorsque l’utilisateur lit un élément qui contient des correspondances d’entités connues, que vous utilisiez ou non une `ItemHasKnownEntity` règle.


- Vous pouvez extraire des chaînes qui sont des entités connues uniquement si les chaînes sont en anglais.
    
- Vous pouvez extraire des entités connues des 2 000 premiers caractères du corps de l’élément, mais pas au-delà. Cette limite de taille permet d’équilibrer le besoin de fonctionnalité et les performances, de sorte qu’Exchange Server et Outlook ne soient pas ralentis par l’analyse et l’identification des instances d’entités connues dans les longs messages et rendez-vous. Notez que cette limite est indépendante du fait que le add-in spécifie ou non une `ItemHasKnownEntity` règle. Si le complément n’utilise pas cette règle, la limite de traitement de règle est celle décrite au point 2 ci-dessous pour les clients riches Outlook.
    
- Vous pouvez extraire des entités à partir de rendez-vous, qui sont des réunions organisées par une personne autre que le propriétaire de la boîte aux lettres. Vous ne pouvez pas extraire d’entités à partir d’éléments de calendrier qui ne sont pas des réunions ou de réunions organisées par le propriétaire de la boîte aux lettres.
    
- Vous pouvez extraire des entités du type à `MeetingSuggestion` partir de messages uniquement, mais pas de rendez-vous.
    
- Vous pouvez extraire des URL qui existent de façon explicite dans le corps d’élément, mais pas des URL incorporées dans un texte de lien hypertexte du corps d’élément HTML. Envisagez plutôt `ItemHasRegularExpressionMatch` d’utiliser une règle pour obtenir des URL explicites et incorporées. Spécifiez en tant que PropertyName et une expression régulière qui correspond aux `BodyAsHTML` URL en tant que _RegExValue_. 
    
- Vous ne pouvez pas extraire des entités à partir d’éléments dans le dossier Éléments envoyés.
    
En outre, les éléments suivants s’appliquent si vous utilisez une règle [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) et peuvent affecter les scénarios dans lequel vous vous attendriez à ce que votre complément soit activé autrement.

- Lors de l’utilisation de la règle, Outlook des chaînes d’entité en anglais uniquement, quels que soient les paramètres régionaux par défaut spécifiés `ItemHasKnownEntity` dans le manifeste.
    
- Lorsque votre application est en cours d’exécution sur un client riche Outlook, attendez-vous à ce que Outlook applique la règle au premier mégaoctet du corps de l’élément et non au reste du corps au-dessus de cette `ItemHasKnownEntity` limite.
    
- Vous ne pouvez pas utiliser une règle pour activer un add-in pour les éléments du `ItemHasKnownEntity` dossier Éléments envoyés.
    

## <a name="see-also"></a>Voir aussi

- [Créer des compléments Outlook pour des formulaires de lecture](read-scenario.md)
- [Extraire des chaînes d’entité d’un élément Outlook](extract-entity-strings-from-an-item.md)
- [Règles d’activation pour les compléments Outlook](activation-rules.md)
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Présentation des autorisations de complément Outlook](understanding-outlook-add-in-permissions.md)
