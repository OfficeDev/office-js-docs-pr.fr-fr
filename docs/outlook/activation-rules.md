---
title: Règles d’activation pour les compléments Outlook
description: Outlook active certains types de complément si le message ou le rendez-vous que l’utilisateur lit ou compose respecte les règles d’activation du complément.
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: a5fc107c27feb5b0535727a42b4d56d21f7dcbc4
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076811"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a>Règles d’activation des compléments contextuels Outlook 

Outlook active certains types de compléments si le message ou le rendez-vous que l’utilisateur lit ou compose respecte les règles d’activation du complément. Cela est vrai pour tous les compléments qui utilisent le schéma de manifeste 1.1. L’utilisateur peut choisir le complément à partir de l’interface utilisateur Outlook afin de le démarrer pour l’élément actuel.

La figure suivante illustre les compléments Outlook activés dans la barre des compléments pour le message dans le volet de lecture.

![Barre de l’application affichant les applications de messagerie en lecture activées.](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a>Spécifier des règles d’activation dans un manifeste


Pour que Outlook activer un complément pour des conditions spécifiques, spécifiez des règles d’activation dans le manifeste du complément à l’aide de l’un des éléments `Rule` suivants :

- [Élément de règle (MailApp complexType)](../reference/manifest/rule.md) : spécifie une règle individuelle.
- [Élément de règle (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) : combine plusieurs règles à l’aide d’opérations logiques.
    

 > [!NOTE]
 > `Rule`L’élément que vous utilisez pour spécifier une règle individuelle est du type complexe [Rule](../reference/manifest/rule.md) abstrait. Chacun des types de règles suivants étend ce `Rule` type complexe abstrait. Ainsi, quand vous spécifiez une règle individuelle dans un manifeste, vous devez utiliser l’attribut [xsi:type](https://www.w3.org/TR/xmlschema-1/) pour définir plus précisément l’un des types de règle suivants.
 > 
 > Par exemple, la règle suivante définit une règle [ItemIs](../reference/manifest/rule.md#itemis-rule) :`<Rule xsi:type="ItemIs" ItemType="Message" />`
 > 
 > `FormType`L’attribut s’applique aux règles d’activation dans le manifeste v1.1, mais n’est pas défini dans `VersionOverrides` la v1.0. Il ne peut donc pas être utilisé lorsque [itemIs](../reference/manifest/rule.md#itemis-rule) est utilisé dans `VersionOverrides` le nœud.

Le tableau suivant répertorie les types de règle disponibles. Vous trouverez plus d’informations dans le tableau et dans les articles indiqués sous [Créer des compléments Outlook pour des formulaires de lecture](read-scenario.md).

<br/>

|**Nom de la règle**|**Formulaires applicables**|**Description**|
|:-----|:-----|:-----|
|[ItemIs](#itemis-rule)|Lecture, composition|Vérifie si l’élément actuel est du type spécifié (message ou rendez-vous). Peut également vérifier la classe de l’élément et le type de formulaire, ainsi qu’éventuellement la classe de message de l’élément.|
|[ItemHasAttachment](#itemhasattachment-rule)|Lecture|Vérifie si l’élément sélectionné contient une pièce jointe.|
|[ItemHasKnownEntity](#itemhasknownentity-rule)|Lecture|Vérifie si l’élément sélectionné contient une ou plusieurs entités reconnues. Plus d’informations : [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md).|
|[ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)|Lecture|Vérifie si l’adresse électronique de l’expéditeur, l’objet et/ou le corps de l’élément sélectionné contient une correspondance avec une expression régulière.Plus d’informations : [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).|
|[RuleCollection](#rulecollection-rule)|Lecture, composition|Associe un ensemble de règles pour vous permettre de former des règles plus complexes.|

## <a name="itemis-rule"></a>Règle ItemIs

Le type complexe **ItemIs** définit une règle qui a pour valeur **true** si l’élément actuel correspond au type d’élément et, éventuellement, la classe de message de l’élément s’il est indiqué dans la règle.

Spécifiez l’un des types d’éléments suivants dans `ItemType` l’attribut d’une **règle ItemIs.** Vous pouvez spécifier plusieurs règles **ItemIs** dans un manifeste. L’élément ItemType simpleType définit les types d’élément Outlook qui prennent en charge les compléments Outlook.

<br/>

|**Value**|**Description**|
|:-----|:-----|
|**Rendez-vous**|Spécifie un élément dans le calendrier Outlook. Par exemple, un élément de réunion auquel une réponse a été donnée et auquel un organisateur et des participants sont associés, ou un rendez-vous auquel n’est associé aucun organisateur ou participant et qui constitue un simple élément de calendrier.Cela correspond à la classe de message IPM.Appointment dans Outlook.|
|**Message**|Spécifie l’un des éléments suivants, généralement reçus dans la boîte de réception : <ul><li><p>Message électronique. Cela correspond à la classe de message IPM.Note dans Outlook.</p></li><li><p>Demande de réunion, réponse à une demande de réunion ou annulation d’une réunion. Cela correspond aux classes de message suivantes dans Outlook :</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|

L’attribut permet de spécifier le mode (lecture ou composition) dans lequel le `FormType` module doit être activé.


 > [!NOTE]
 > L’attribut ItemIs `FormType` est défini dans les schémas v1.1 et ultérieur, mais pas dans la `VersionOverrides` v1.0. N’incluez pas `FormType` l’attribut lors de la définition des commandes de add-in.

Une fois qu’un complément est activé, vous pouvez utiliser la propriété [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) pour obtenir l’élément actuellement sélectionné dans Outlook, et la propriété [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) pour obtenir le type de l’élément actuel.

Vous pouvez éventuellement utiliser l’attribut pour spécifier la classe de message de l’élément et l’attribut pour spécifier si la règle doit être true lorsque l’élément est une sous-classe de la classe `ItemClass` `IncludeSubClasses` spécifiée. 

Pour plus d’informations sur les classes de message, reportez-vous à la rubrique [Types d’éléments et classes de messages](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).

L’exemple suivant illustre une règle **ItemIs** permettant aux utilisateurs d’afficher le complément dans la barre de compléments Outlook lorsqu’ils lisent un message :

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

L’exemple suivant illustre une règle **ItemIs** permettant aux utilisateurs d’afficher le complément dans la barre de compléments Outlook lorsqu’ils lisent un message ou un rendez-vous.

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a>Règle ItemHasAttachment


Le `ItemHasAttachment` type complexe définit une règle qui vérifie si l’élément sélectionné contient une pièce jointe.

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a>Règle ItemHasKnownEntity

Avant qu’un élément ne soit mis à la disposition d’un complément, le serveur l’examine pour déterminer si l’objet ou le corps contient du texte susceptible de correspondre à l’une des entités connues. Si l’une de ces entités est trouvée, elle est placée dans une collection d’entités connues accessibles à l’aide de la ou de la `getEntities` `getEntitiesByType` méthode de cet élément.

Vous pouvez spécifier une règle à l’aide de celle qui affiche votre add-in lorsqu’une entité du type spécifié `ItemHasKnownEntity` est présente dans l’élément. Vous pouvez spécifier les entités connues suivantes dans `EntityType` l’attribut d’une `ItemHasKnownEntity` règle :

- Address
- Contact
- EmailAddress
- MeetingSuggestion
- PhoneNumber
- TaskSuggestion
- URL
    
Vous pouvez éventuellement inclure une expression régulière dans l’attribut afin que votre add-in s’affiche uniquement lorsqu’une entité correspond à `RegularExpression` l’expression régulière présente. Pour obtenir des correspondances avec des expressions régulières spécifiées dans les règles, vous pouvez utiliser la ou la méthode pour l’élément `ItemHasKnownEntity` `getRegExMatches` Outlook actuellement `getFilteredEntitiesByName` sélectionné.

L’exemple suivant montre une collection d’éléments qui montrent le add-in lorsque l’une des entités connues spécifiées est présente `Rule` dans le message.

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

L’exemple suivant montre une règle avec un attribut qui active le module lorsqu’une URL contenant le `ItemHasKnownEntity` mot « contoso » est présente `RegularExpression` dans un message.


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

Pour plus d’informations sur les entités dans les règles d’activation, voir [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md).


## <a name="itemhasregularexpressionmatch-rule"></a>Règle ItemHasRegularExpressionMatch

Le type complexe définit une règle qui utilise une expression régulière pour correspondre au contenu de la propriété spécifiée `ItemHasRegularExpressionMatch` d’un élément. Si le texte correspondant à l’expression régulière est trouvé dans la propriété spécifiée de l’élément, Outlook active la barre de compléments et affiche le complément. Vous pouvez utiliser la ou la méthode de l’objet qui représente l’élément actuellement sélectionné pour obtenir des `getRegExMatches` `getRegExMatchesByName` correspondances pour l’expression régulière spécifiée.

L’exemple suivant montre un exemple qui active le add-in lorsque le corps de l’élément sélectionné contient « apple », « premier » ou « coco » en ignorant la `ItemHasRegularExpressionMatch` cas.

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

Pour plus d’informations sur l’utilisation de la règle, voir Utiliser des règles d’activation d’expression régulière pour afficher `ItemHasRegularExpressionMatch` [Outlook complément.](use-regular-expressions-to-show-an-outlook-add-in.md)


## <a name="rulecollection-rule"></a>Règle RuleCollection


Le `RuleCollection` type complexe combine plusieurs règles en une seule règle. Vous pouvez spécifier si les règles de la collection doivent être combinées avec un OU logique ou un AND logique à l’aide de `Mode` l’attribut.

Quand un ET logique est spécifié, un élément doit correspondre à toutes les règles spécifiées dans le regroupement pour permettre l’affichage du complément. Quand un OU logique est spécifié, tout élément qui correspond à l’une des règles spécifiées dans le regroupement permet l’affichage du complément.

Vous pouvez combiner `RuleCollection` des règles pour former des règles complexes. L’exemple suivant illustre l’activation du complément lorsque l’utilisateur affiche un élément de rendez-vous ou de message et que l’objet ou le corps de l’élément contient une adresse.

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

L’exemple suivant illustre l’activation du complément lorsque l’utilisateur compose un message ou affiche un rendez-vous, et que l’objet ou le corps du rendez-vous contient une adresse.

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a>Limites pour les règles et les expressions régulières


Pour fournir une expérience satisfaisante avec les compléments Outlook, vous devez vous conformer aux directives d’activation et d’utilisation des API. Le tableau suivant indique les limites générales pour les expressions régulières et les règles, mais il existe des règles spécifiques pour différentes applications. Pour plus d’informations, voir [Limites d’activation et d’API JavaScript des compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) et [Résoudre les problèmes d’activation des compléments Outlook](troubleshoot-outlook-add-in-activation.md).

<br/>

|**Élément de complément**|**Conseils**|
|:-----|:-----|
|Taille de manifeste|Inférieur à 256 Ko.|
|Règles|Pas plus de 15 règles.|
|ItemHasKnownEntity|Un riche client Outlook appliquera la règle au premier mégaoctet du corps, mais pas au reste.|
|Expressions régulières|Pour les règles ItemHasKnownEntity ou ItemHasRegularExpressionMatch pour toutes Outlook applications :<br><ul><li>Ne spécifiez pas plus de 5 expressions régulières dans les règles d’activation pour un complément Outlook. Vous ne pouvez pas installer de complément si vous dépassez cette limite.</li><li>Spécifiez des expressions régulières dont les résultats sont renvoyés par l’appel de la méthode <b>getRegExMatches</b> dans les 50 premières correspondances. </li><li>Spécifiez des assertions avant dans les expressions régulières, mais pas d’assertions arrière, `(?<=text)`, ni d’assertions arrière négatives `(?<!text)`.</li><li>Spécifiez des expressions régulières dont la correspondance ne dépasse pas les limites figurant dans le tableau ci-dessous.<br/><br/><table><tr><th>Limite de longueur d’une correspondance d’expression régulière</th><th>Clients riches Outlook</th><th>Outlook sur iOS et Android</th></tr><tr><td>Corps d’élément en texte brut</td><td>1,5 Ko</td><td>3 Ko</td></tr><tr><td>Corps d’élément en HTML</td><td>3 Ko</td><td>3 Ko</td></tr></table>|

## <a name="see-also"></a>Voir aussi

- [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md)
    
