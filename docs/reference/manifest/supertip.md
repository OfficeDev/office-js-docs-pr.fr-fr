# <a name="supertip"></a>Supertip

Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles [Bouton](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Title](#title)        | Oui |   Texte du supertip.         |
|  [Description](#description)  | Oui |  Description du supertip.    |

### <a name="title"></a>Titre

Obligatoire. Texte du superTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** dans l’élément [Resources](resources.md).

### <a name="description"></a>Description

Obligatoire. Description du superTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **LongStrings** dans l’élément [Resources](resources.md).

## <a name="example"></a>Exemple

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
