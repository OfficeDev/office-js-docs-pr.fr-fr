# <a name="page-element"></a>Élément Page

Définit les paramètres de la page HTML utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

Aucun

## <a name="child-elements"></a>Éléments enfants

|  Élément  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne contenant l’ID de ressource du fichier HTML utilisé par les fonctions personnalisées. |

## <a name="example"></a>Exemple

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
