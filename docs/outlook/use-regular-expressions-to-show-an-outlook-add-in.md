---
title: Utiliser les règles d’activation d’expression régulière afin d’afficher un complément
description: Découvrez comment utiliser les règles d’activation d’expression régulière pour les compléments contextuels Outlook.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: ed2fbbfcf7bf55e04f4ec6f225e29fb43ec99639
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467089"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a>Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook

Vous pouvez spécifier des règles d’expressions régulières pour qu’un [complément contextuel](contextual-outlook-add-ins.md) soit activé lorsqu’une correspondance est trouvée dans les champs spécifiques du message. Les compléments contextuels s’activent uniquement en mode lecture. Outlook n’active pas les compléments contextuels lorsque l’utilisateur compose un élément. Il existe également d’autres scénarios où Outlook n’active pas les compléments, par exemple, les éléments signés numériquement. Pour plus d’informations, reportez-vous à la rubrique [Règles d’activation pour les compléments Outlook](activation-rules.md).

[!include[JSON manifest does not support contextual add-ins](../includes/json-manifest-outlook-contextual-not-supported.md)]

Vous pouvez spécifier une expression régulière dans le cadre d’une règle [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) ou [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) dans le manifeste XML du complément. Les règles sont spécifiées dans un point d’extension [DetectedEntity](/javascript/api/manifest/extensionpoint#detectedentity).

Outlook évalue les expressions régulières en fonction des règles définies pour l’interpréteur JavaScript utilisé par le navigateur de l’ordinateur client. Outlook prend en charge la même liste de caractères spéciaux que tous les processeurs XML. Le tableau suivant répertorie ces caractères spéciaux. Vous pouvez utiliser ces caractères dans une expression régulière en spécifiant la séquence d’échappement du caractère correspondant, comme décrit dans le tableau suivant.

|Caractère|Description|Séquence d’échappement à utiliser|
|:-----|:-----|:-----|
|`"`|Guillemets doubles|`&quot;`|
|`&`|Esperluette|`&amp;`|
|`'`|Apostrophe|`&apos;`|
|`<`|Signe inférieur à|`&lt;`|
|`>`|Signe supérieur à|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a>Règle ItemHasRegularExpressionMatch

Une règle `ItemHasRegularExpressionMatch` est utile dans le contrôle de l’activation d’un complément basé sur les valeurs spécifiques d’une propriété prise en charge. La règle `ItemHasRegularExpressionMatch` contient les attributs suivants.

|Nom de l’attribut|Description|
|:-----|:-----|
|`RegExName`|Spécifie le nom de l’expression régulière afin que vous puissiez vous référer à l’expression dans le code de votre complément.|
|`RegExValue`|Spécifie l’expression régulière qui sera évaluée pour déterminer si le complément doit être affiché.|
|`PropertyName`|Spécifie le nom de la propriété par rapport à laquelle l’expression sera évaluée. Les valeurs autorisées sont `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress` et `Subject`.<br/><br/>Si vous spécifiez `BodyAsHTML`, Outlook applique seulement l’expression régulière si le corps de l’élément est du code HTML. Si ce n’est pas le cas, Outlook ne renvoie aucune correspondance pour cette expression régulière.<br/><br/>Si vous spécifiez `BodyAsPlaintext`, Outlook applique toujours l’expression régulière au corps de l’élément.<br/><br/>**Important:** Si vous devez spécifier l’attribut **Highlight** pour l’élément **\<Rule\>** , vous devez définir l’attribut **PropertyName sur** `BodyAsPlaintext`. |
|`IgnoreCase`|Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par `RegExName`.|
| `Highlight` | Spécifie la façon dont le client doit mettre en évidence le texte correspondant. Cet élément peut uniquement s’appliquer à des éléments `Rule` au sein d’éléments `ExtensionPoint`. Peut correspondre à l’une des valeurs suivantes : `all` ou `none`. Si non spécifié, la valeur par défaut est `all`.<br/><br/>**Important:** Pour spécifier l’attribut **Highlight** dans l’élément **\<Rule\>**, vous devez définir l’attribut `BodyAsPlaintext`**PropertyName sur** . |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a>Meilleures pratiques pour l’utilisation d’expressions régulières dans les règles

Accordez une attention particulière aux éléments suivants lorsque vous utilisez des expressions régulières.

- Si vous spécifiez une `ItemHasRegularExpressionMatch` règle sur le corps d’un élément, l’expression régulière doit filtrer davantage le corps et ne pas tenter de retourner l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` la tentative d’obtention de l’intégralité du corps d’un élément ne retourne pas toujours les résultats attendus.
- Le corps en texte brut renvoyé sur un navigateur peut être légèrement différent sur un autre. Si vous utilisez une règle `ItemHasRegularExpressionMatch` avec `BodyAsPlaintext` comme attribut `PropertyName`, testez votre expression régulière sur tous les navigateurs pris en charge par votre complément.

    Comme différents navigateurs utilisent diverses méthodes pour obtenir le corps du texte d’un élément sélectionné, vous devez vous assurer que votre expression régulière prend en charge les fines différences qui peuvent être renvoyées dans le cadre du corps de texte. Par exemple, certains navigateurs, comme Internet Explorer 9, utilisent la propriété `innerText` du DOM, tandis que d’autres, comme Firefox, utilisent la méthode `.textContent()` afin d’obtenir le corps du texte d’un élément. En outre, différents navigateurs peuvent renvoyer des sauts de ligne de manière différente : un saut de ligne correspond à `\r\n` sur Internet Explorer, et `\n` dans Firefox et Chrome. Pour plus d’informations, consultez la page sur la [compatibilité DOM W3C - HTML](https://quirksmode.org/dom/html/).

- Le corps HTML d’un élément est légèrement différent entre un client riche Outlook et Outlook sur le web ou Outlook Mobile. Définissez attentivement vos expressions régulières.

- Selon le client Outlook, le type d’appareil ou la propriété sur lequel une expression régulière est appliquée, il existe d’autres meilleures pratiques et limites pour chacun des clients que vous devez connaître lors de la conception d’expressions régulières en tant que règles d’activation. Pour plus d’informations, voir [Limites d’activation et d’API JavaScript des compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).

### <a name="examples"></a>Exemples

La règle `ItemHasRegularExpressionMatch` suivante active le complément chaque fois que l’adresse de messagerie SMTP de l’expéditeur correspond à `@contoso`, indépendamment des caractères majuscules et minuscules.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

L’exemple suivant montre une autre manière de spécifier la même expression régulière à l’aide de l’attribut `IgnoreCase`.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

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
> Outlook peut extraire uniquement des chaînes d’entité en anglais, indépendamment des paramètres régionaux par défaut spécifiés dans le manifeste. Seuls les messages prennent en charge le `MeetingSuggestion` type d’entité ; les rendez-vous ne le prennent pas en charge. Vous ne pouvez pas extraire d’entités d’éléments du dossier **Éléments envoyés** , ni utiliser une `ItemHasKnownEntity` règle pour activer un complément pour les éléments du dossier **Éléments envoyés** .

La règle `ItemHasKnownEntity` prend en charge les attributs dans le tableau suivant. Notez que, bien que la spécification d’une expression régulière soit facultative dans une règle `ItemHasKnownEntity`, si vous choisissez d’utiliser une expression régulière comme filtre d’entité, vous devez spécifier à la fois l’attribut `RegExFilter` et `FilterName`.

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

Vous pouvez obtenir des correspondances avec une expression régulière à l’aide des méthodes suivantes sur l’élément actif.

- [getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) renvoie les correspondances dans l’élément actuel pour toutes les expressions régulières spécifiées dans les règles `ItemHasRegularExpressionMatch` et `ItemHasKnownEntity` du complément.

- [getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) renvoie les correspondances dans l’élément actuel pour l’expression régulière identifiée, spécifiée dans une règle `ItemHasRegularExpressionMatch` du complément.

- [getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) renvoie les instances complètes des entités qui contiennent des correspondances avec l’expression régulière identifiée, spécifiée dans une règle `ItemHasKnownEntity` du complément.

Lorsque les expressions régulières sont évaluées, les correspondances sont renvoyées vers votre complément dans un objet tableau. Pour `getRegExMatches`, cet objet a un identifiant correspondant au nom de l’expression régulière.

> [!NOTE]
> Outlook ne retourne pas de correspondances dans un ordre particulier dans le tableau. En outre, vous ne devez pas supposer que les correspondances sont retournées dans le même ordre dans ce tableau, même lorsque vous exécutez le même complément sur chacun de ces clients sur le même élément dans la même boîte aux lettres.

### <a name="examples"></a>Exemples

L’exemple suivant montre un regroupement de règles qui contient une règle `ItemHasRegularExpressionMatch` avec une expression régulière nommée `videoURL`.

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

L’exemple suivant utilise `getRegExMatches` dans l’élément actuel pour définir une variable `videos` pour les résultats de la règle `ItemHasRegularExpressionMatch` précédente.

```js
const videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

Multiple matches are stored as array elements in that object. The following code example shows how to iterate over the matches for a regular expression named  `reg1` to build a string to display as HTML.

```js
function initDialer()
{
    let myEntities;
    let myString;
    let myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (let i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

Voici un exemple de règle `ItemHasKnownEntity` qui spécifie l’entité `MeetingSuggestion` et une expression régulière nommée `CampSuggestion`. Outlook active le complément s’il détecte que l’élément sélectionné contient une suggestion de réunion, et que l’objet ou le corps contient le terme `WonderCamp`.

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

L’exemple de code suivant utilise `getFilteredEntitiesByName` sur l’élément actuel pour définir une variable `suggestions` pour un tableau des suggestions de réunion détectées pour la règle `ItemHasKnownEntity` précédente.

```js
const suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a>Voir aussi

- [Complément Outlook : numéro de commande Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - Exemple de complément contextuel qui est activé en fonction d’une correspondance d’expression régulière.
- [Créer des compléments Outlook pour des formulaires de lecture](read-scenario.md)
- [Règles d’activation pour les compléments Outlook](activation-rules.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md)
- [Meilleures pratiques pour les expressions régulières dans le .NET Framework](/dotnet/standard/base-types/best-practices)
