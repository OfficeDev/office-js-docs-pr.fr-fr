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
| **FormType** | Non (dans [ExtensionPoint](extensionpoint.md)), Oui (dans [OfficeApp](officeapp.md)) | Spécifie si l’application doit apparaître dans le formulaire de lecture ou de modification pour l’élément. Peut correspondre à l’une des valeurs suivantes : `Read`, `Edit`, `ReadOrEdit`. Si spécifiée dans un `Rule` dans un `ExtensionPoint`, cette valeur DOIT être `Read`. |
| **ItemClass** | Non | Spécifie la classe de message personnalisé à mettre en correspondance. Pour plus d’informations, voir l’article relatif à l’[Activation d’un complément de messagerie dans Outlook pour une classe de message spécifique](https://docs.microsoft.com/outlook/add-ins/activation-rules). |
| **IncludeSubClasses** | Non | Spécifie si la règle doit donner la valeur true si l’élément est une sous-classe de la classe de message spécifiée. Par défaut, la valeur est `false`. |

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
| **IgnoreCase** | Non | Indique d’ignorer la casse lors de l’exécution de l’expression régulière spécifiée par l’attribut **RegExFilter**. |
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
| **PropertyName** | Oui | Spécifie le nom de la propriété par rapport à laquelle l’expression sera évaluée. Peut correspondre à l’une des valeurs suivantes : `Subject`, `BodyAsPlaintext`, `BodyAsHtml` ou `SenderSTMPAddress`. |
| **IgnoreCase** | Non | Indique d’ignorer la casse lors de l’exécution de l’expression régulière. |
| **Highlight** | Non | **Remarque :** cela s’applique uniquement aux éléments **Rule** au sein des éléments **ExtensionPoint**. Spécifie comment le client doit mettre en surbrillance le texte correspondant. Peut correspondre à l’une des valeurs suivantes : `all` ou `none`. Si non spécifié, la valeur par défaut est `all`. |

### <a name="example"></a>Exemple

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## <a name="rulecollection"></a>RuleCollection

Définit une collection de règles et l’opérateur logique à utiliser lors de leur évaluation.

### <a name="attributes"></a>Attributs

| Attribut | Obligatoire | Description |
|:-----|:-----|:-----|
| **Mode** | Oui | Spécifie l’opérateur logique à utiliser lors de l’évaluation de cette collection de règles. Peut être `And` ou `Or`. |

### <a name="example"></a>Exemple

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a>Voir aussi

- [Règles d’activation pour les compléments Outlook](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)