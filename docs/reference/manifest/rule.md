---
title: Élément Rule dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 07037c43c111f735a7354a048066e4c4a88f7637
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450470"
---
# <a name="rule-element"></a>Élément Rule

Spécifie les règles d’activation à évaluer pour ce complément de messagerie contextuel.

**Type de complément :** complément de messagerie contextuel

## <a name="contained-in"></a>Contenu dans

- [OfficeApp](officeapp.md)
- [ExtensionPoint](extensionpoint.md)

## <a name="attributes"></a>Attributs

| Attribut | Obligatoire | Description |
|:-----|:-----|:-----|
| **xsi:type** | Oui | Type de règle en cours de définition. |

Le type de règle peut correspondre à l’une des valeurs suivantes.

- [ItemIs](#itemis-rule)
- [ItemHasAttachment](#itemhasattachment-rule)
- [ItemHasKnownEntity](#itemhasknownentity-rule)
- [ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)
- [RuleCollection](#rulecollection)

## <a name="itemis-rule"></a>Règle ItemIs

Définit une règle qui donne la valeur true si l’élément sélectionné est du type spécifié.

### <a name="attributes"></a>Attributs

| Attribut | Obligatoire | Description |
|:-----|:-----|:-----|
| **ItemType** | Oui | Spécifie le type d’élément à mettre en correspondance. Peut être `Message` ou `Appointment`. Le type d’élément `Message` inclut e-mails, demandes de réunion, réponses à une demande de réunion et annulations de réunion. |
| **FormType** | Non (dans [ExtensionPoint](extensionpoint.md)), Oui (dans [App_office](officeapp.md)) | Spécifie si l’application doit apparaître dans le formulaire de lecture ou de modification pour l’élément. Peut correspondre à l’une des valeurs suivantes : `Read`, `Edit`, `ReadOrEdit`. Si spécifiée dans un `Rule` dans un `ExtensionPoint`, cette valeur DOIT être `Read`. |
| **ItemClass** | Non | Spécifie la classe de message personnalisé à mettre en correspondance. Pour plus d’informations, voir l’article relatif à l’[activation d’un complément de messagerie dans Outlook pour une classe de message spécifique](/outlook/add-ins/activation-rules). |
| **IncludeSubClasses** | Non | Spécifie si la règle doit donner la valeur true si l’élément est une sous-classe de la classe de message spécifiée ; par défaut, la valeur est `false`. |

### <a name="example"></a>Exemple

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a>Règle ItemHasAttachment

Définit une règle qui donne la valeur true si l’élément contient une pièce jointe.

### <a name="example"></a>Exemple

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a>Règle ItemHasKnownEntity

Définit une règle qui donne la valeur true si l’élément contient dans son objet ou son corps du texte correspondant au type d’entité spécifié.

### <a name="attributes"></a>Attributs

| Attribut | Obligatoire | Description |
|:-----|:-----|:-----|
| **EntityType** | Oui | Spécifie le type d’entité à rechercher pour que la règle donne la valeur true. Peut correspondre à l’une des valeurs suivantes : `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` ou `Contact`. |
| **RegExFilter** | Non | Spécifie une expression régulière à exécuter par rapport à cette entité à des fins d’activation. |
| **FilterName** | Non | Spécifie le nom du filtre d’expression régulière, afin qu’il soit possible par la suite de s’y référer dans le code de votre complément. |
| **IgnoreCase** | Non | Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par l’attribut **RegExFilter**. |
| **Highlight** | Non | **Remarque :** cela s’applique uniquement aux éléments **Rule** au sein des éléments **ExtensionPoint**. Spécifie comment le client doit mettre en surbrillance les entités correspondantes. Peut correspondre à l’une des valeurs suivantes : `all` ou `none`. Si non spécifié, la valeur par défaut est `all`. |

### <a name="example"></a>Exemple

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a>Règle ItemHasRegularExpressionMatch

Définit une règle qui donne la valeur true si une correspondance de l’expression régulière spécifiée est trouvée dans la propriété spécifiée de l’élément.

### <a name="attributes"></a>Attributs

| Attribut | Obligatoire | Description |
|:-----|:-----|:-----|
| **RegExName** | Oui | Spécifie le nom de l’expression régulière afin que vous puissiez vous référer à l’expression dans le code de votre complément. |
| **RegExValue** | Oui | Spécifie l’expression régulière qui sera évaluée pour déterminer si le complément de messagerie doit être affiché. |
| **PropertyName** | Oui | Spécifie le nom de la propriété par rapport à laquelle l’expression sera évaluée. Les options disponibles sont les suivantes : `Subject`, `BodyAsPlaintext`, `BodyAsHTML` ou `SenderSMTPAddress`.<br/><br/>Si vous spécifiez `BodyAsHTML`, Outlook applique seulement l’expression régulière si le corps de l’élément est du code HTML. Si ce n’est pas le cas, Outlook ne renvoie aucune correspondance pour cette expression régulière.<br/><br/>Si vous spécifiez `BodyAsPlaintext`, Outlook applique toujours l’expression régulière au corps de l’élément.<br/><br/>**Remarque :** vous devez donner la valeur `BodyAsPlaintext` à l’attribut **PropertyName** si vous spécifiez l’attribut **Highlight** pour l’élément **Rule**.|
| **IgnoreCase** | Non | Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par l’attribut **RegExName**. |
| **Highlight** | Non | Spécifie comment le client doit mettre en surbrillance le texte correspondant. Cet attribut ne peut être appliqué qu’aux éléments **Rule** au sein des éléments **ExtensionPoint**. Peut correspondre à l’une des valeurs suivantes : `all` ou `none`. Si non spécifié, la valeur par défaut est `all`.<br/><br/>**Remarque :** vous devez donner la valeur `BodyAsPlaintext` à l’attribut **PropertyName** si vous spécifiez l’attribut **Highlight** pour l’élément **Rule**.
|

### <a name="example"></a>Exemple

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a>RuleCollection

Définit une collection de règles et l’opérateur logique à utiliser lors de leur évaluation.

### <a name="attributes"></a>Attributs

| Attribut | Obligatoire | Description |
|:-----|:-----|:-----|
| **Mode** | Oui | Spécifie l’opérateur logique à utiliser lors de l’évaluation de cette collection de règles. Il peut s’agir des éléments `And` ou `Or`. |

### <a name="example"></a>Exemple

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a>Voir aussi

- [Règles d’activation pour les compléments Outlook](/outlook/add-ins/activation-rules)
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)
