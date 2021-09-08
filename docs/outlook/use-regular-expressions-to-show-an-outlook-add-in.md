---
title: Utiliser les règles d’activation d’expression régulière afin d’afficher un complément
description: Découvrez comment utiliser les règles d’activation d’expression régulière pour les compléments contextuels Outlook.
ms.date: 07/28/2020
localization_priority: Normal
ms.openlocfilehash: d334ba6b2e0f044fc8d876cd6edd218743ccb390
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938864"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a>Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook

Vous pouvez spécifier des règles d’expressions régulières pour qu’un [complément contextuel](contextual-outlook-add-ins.md) soit activé lorsqu’une correspondance est trouvée dans les champs spécifiques du message. Les compléments contextuels sont activés uniquement en mode lecture. Outlook n’active pas de compléments contextuels lorsque l’utilisateur compose un élément. Il existe également d’autres scénarios dans Outlook n’active pas les modules, par exemple, les éléments signés numériquement. Pour plus d’informations, reportez-vous à la rubrique [Règles d’activation pour les compléments Outlook](activation-rules.md).

Vous pouvez spécifier une expression régulière dans le cadre d’une règle [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) ou [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) dans le manifeste XML du complément. Les règles sont spécifiées dans un point d’extension [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity).

Outlook évalue les expressions régulières en fonction des règles définies pour l’interpréteur JavaScript utilisé par le navigateur de l’ordinateur client. Outlook prend en charge la même liste de caractères spéciaux que tous les processeurs XML. Le tableau suivant répertorie ces caractères spéciaux. Vous pouvez les utiliser dans une expression régulière en spécifiant la séquence d’échappement pour le caractère correspondant, comme décrit dans le tableau suivant.

<br/>

|Caractère|Description|Séquence d’échappement à utiliser|
|:-----|:-----|:-----|
|`"`|Guillemets doubles|`&quot;`|
|`&`|Esperluette|`&amp;`|
|`'`|Apostrophe|`&apos;`|
|`<`|Signe inférieur à|`&lt;`|
|`>`|Signe supérieur à|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a>Règle ItemHasRegularExpressionMatch

Une règle `ItemHasRegularExpressionMatch` est utile dans le contrôle de l’activation d’un complément basé sur les valeurs spécifiques d’une propriété prise en charge. La règle `ItemHasRegularExpressionMatch` contient les attributs suivants.

<br/>

|Nom de l’attribut|Description|
|:-----|:-----|
|`RegExName`|Spécifie le nom de l’expression régulière afin que vous puissiez vous référer à l’expression dans le code de votre complément.|
|`RegExValue`|Spécifie l’expression régulière qui sera évaluée pour déterminer si le complément doit être affiché.|
|`PropertyName`|Spécifie le nom de la propriété par rapport à laquelle l’expression sera évaluée. Les valeurs autorisées sont `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress` et `Subject`.<br/><br/>Si vous spécifiez `BodyAsHTML`, Outlook applique seulement l’expression régulière si le corps de l’élément est du code HTML. Si ce n’est pas le cas, Outlook ne renvoie aucune correspondance pour cette expression régulière.<br/><br/>Si vous spécifiez `BodyAsPlaintext`, Outlook applique toujours l’expression régulière au corps de l’élément.<br/><br/>**Remarque :** vous devez définir l’attribut `PropertyName` sur `BodyAsPlaintext` si vous spécifiez l’attribut `Highlight` pour l’élément `Rule`.|
|`IgnoreCase`|Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par `RegExName`.|
| `Highlight` | Spécifie la façon dont le client doit mettre en évidence le texte correspondant. Cet élément peut uniquement s’appliquer à des éléments `Rule` au sein d’éléments `ExtensionPoint`. Peut correspondre à l’une des valeurs suivantes : `all` ou `none`. Si non spécifié, la valeur par défaut est `all`.<br/><br/>**Remarque :** vous devez définir l’attribut `PropertyName` sur `BodyAsPlaintext` si vous spécifiez l’attribut `Highlight` pour l’élément `Rule`. |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a>Meilleures pratiques pour l’utilisation d’expressions régulières dans les règles

Prêtez une attention particulière aux questions suivantes lorsque vous utilisez des expressions régulières.

- Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour le corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour essayer d’obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.
- Le corps en texte brut renvoyé sur un navigateur peut être légèrement différent sur un autre. Si vous utilisez une règle `ItemHasRegularExpressionMatch` avec `BodyAsPlaintext` comme attribut `PropertyName`, testez votre expression régulière sur tous les navigateurs pris en charge par votre complément.

    Comme différents navigateurs utilisent diverses méthodes pour obtenir le corps du texte d’un élément sélectionné, vous devez vous assurer que votre expression régulière prend en charge les fines différences qui peuvent être renvoyées dans le cadre du corps de texte. Par exemple, certains navigateurs, comme Internet Explorer 9, utilisent la propriété `innerText` du DOM, tandis que d’autres, comme Firefox, utilisent la méthode `.textContent()` afin d’obtenir le corps du texte d’un élément. En outre, différents navigateurs peuvent renvoyer des sauts de ligne de manière différente : un saut de ligne correspond à `\r\n` sur Internet Explorer, et `\n` dans Firefox et Chrome. Pour plus d’informations, consultez la page sur la [compatibilité DOM W3C - HTML](https://quirksmode.org/dom/html/).

- Le corps HTML d’un élément est légèrement différent entre un client riche Outlook et Outlook sur le web ou Outlook Mobile. Définissez attentivement vos expressions régulières.

- Selon le client Outlook, le type d’appareil ou la propriété sur qui une expression régulière est appliquée, il existe d’autres meilleures pratiques et limites pour chacun des clients que vous devez connaître lors de la conception d’expressions régulières en tant que règles d’activation. Pour plus d’informations, voir [Limites d’activation et d’API JavaScript des compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).

### <a name="examples"></a>Exemples

La règle `ItemHasRegularExpressionMatch` suivante active le complément chaque fois que l’adresse de messagerie SMTP de l’expéditeur correspond à `@contoso`, indépendamment des caractères majuscules et minuscules.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

<br/>

L’exemple suivant montre une autre manière de spécifier la même expression régulière à l’aide de l’attribut `IgnoreCase`.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

<br/>

La règle `ItemHasRegularExpressionMatch` suivante active le complément chaque fois qu’un symbole de valeur est inclus dans le corps de l’élément actuel.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a>Règle ItemHasKnownEntity

Une règle `ItemHasKnownEntity` active un complément en fonction de l'existence d'une entité dans le sujet ou le corps de l'élément sélectionné. Le type [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) définit les entités prises en charge. L’application d’une expression régulière sur une règle `ItemHasKnownEntity` convient lorsque l’activation est basée sur un sous-ensemble de valeurs pour une entité (par exemple, un ensemble spécifique d’URL, ou des numéros de téléphone avec un certain code régional).

> [!NOTE]
> Outlook peut extraire uniquement des chaînes d’entité en anglais, indépendamment des paramètres régionaux par défaut spécifiés dans le manifeste. Seuls les messages prennent en charge le type d’entité `MeetingSuggestion`. Ce n’est pas le cas des rendez-vous. Vous ne pouvez pas extraire les entités des éléments figurant dans le dossier **Éléments envoyés**, ni utiliser une règle `ItemHasKnownEntity` afin d’activer un complément pour les éléments du dossier **Éléments envoyés**.

La règle `ItemHasKnownEntity` prend en charge les attributs dans le tableau suivant. Notez que, bien que la spécification d’une expression régulière soit facultative dans une règle `ItemHasKnownEntity`, si vous choisissez d’utiliser une expression régulière comme filtre d’entité, vous devez spécifier à la fois l’attribut `RegExFilter` et `FilterName`.

<br/>

|Nom de l’attribut|Description|
|:-----|:-----|
|`EntityType`|Spécifie le type d’entité à rechercher pour que la règle donne la valeur `true`. Utilisez plusieurs règles pour spécifier plusieurs types d’entités.|
|`RegExFilter`|Spécifie une expression régulière qui filtre les instances de l’entité spécifiée par `EntityType`.|
|`FilterName`|Spécifie le nom de l’expression régulière spécifiée par `RegExFilter`, afin qu’il soit possible d’y faire référence ultérieurement par code.|
|`IgnoreCase`|Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par `RegExFilter`.|

### <a name="examples"></a>Exemples

La règle `ItemHasKnownEntity` suivante active le complément chaque fois qu’une URL se trouve dans l’objet ou le corps de l’élément actuel, et qu’elle contient la chaîne `youtube`, indépendamment de la casse de cette chaîne.

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a>Utilisation des résultats d’expressions régulières dans le code

Vous pouvez obtenir des correspondances avec une expression régulière en utilisant les méthodes suivantes sur l’élément actuel.

- [getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) renvoie les correspondances dans l’élément actuel pour toutes les expressions régulières spécifiées dans les règles `ItemHasRegularExpressionMatch` et `ItemHasKnownEntity` du complément.

- [getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) renvoie les correspondances dans l’élément actuel pour l’expression régulière identifiée, spécifiée dans une règle `ItemHasRegularExpressionMatch` du complément.

- [getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) renvoie les instances complètes des entités qui contiennent des correspondances avec l’expression régulière identifiée, spécifiée dans une règle `ItemHasKnownEntity` du complément.

Lorsque les expressions régulières sont évaluées, les correspondances sont renvoyées vers votre complément dans un objet tableau. Pour `getRegExMatches`, cet objet a un identifiant correspondant au nom de l’expression régulière.

> [!NOTE]
> Les correspondances renvoyées par Outlook ne sont pas classées dans un ordre particulier dans le tableau. Par ailleurs, vous ne devez pas supposer que les correspondances sont renvoyées dans le même ordre dans ce tableau, même lorsque vous exécutez le même complément sur chacun de ces clients sur le même élément de la même boîte aux lettres.

### <a name="examples"></a>Exemples

L’exemple suivant montre un regroupement de règles qui contient une règle `ItemHasRegularExpressionMatch` avec une expression régulière nommée `videoURL`.

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

<br/>

L’exemple suivant utilise `getRegExMatches` dans l’élément actuel pour définir une variable `videos` pour les résultats de la règle `ItemHasRegularExpressionMatch` précédente.

```js
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

<br/>

Plusieurs correspondances sont stockées comme éléments d’un tableau dans cet objet. L’exemple de code suivant montre comment réaliser une itération sur les correspondances pour une expression régulière nommée  `reg1` pour construire une chaîne à afficher sous la forme HTML.

```js
function initDialer()
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

<br/>

Voici un exemple de règle `ItemHasKnownEntity` qui spécifie l’entité `MeetingSuggestion` et une expression régulière nommée `CampSuggestion`. Outlook active le complément s’il détecte que l’élément sélectionné contient une suggestion de réunion, et que l’objet ou le corps contient le terme `WonderCamp`.

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

<br/>

L’exemple de code suivant utilise `getFilteredEntitiesByName` sur l’élément actuel pour définir une variable `suggestions` pour un tableau des suggestions de réunion détectées pour la règle `ItemHasKnownEntity` précédente.

```js
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a>Voir aussi

- [Complément Outlook : numéro de commande Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - Exemple de complément contextuel qui est activé en fonction d’une correspondance d’expression régulière.
- [Créer des compléments Outlook pour des formulaires de lecture](read-scenario.md)
- [Règles d’activation pour les compléments Outlook](activation-rules.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md)
- [Meilleures pratiques pour les expressions régulières dans .NET Framework](/dotnet/standard/base-types/best-practices)
