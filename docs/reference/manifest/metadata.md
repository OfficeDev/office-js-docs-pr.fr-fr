# <a name="metadata-element"></a>Élément Metadata

Définit les paramètres de métadonnées utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

Aucun

## <a name="child-elements"></a>Éléments enfants

|  Élément  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne contenant l’ID de ressource du fichier JSON utilisé par les fonctions personnalisées. |

## <a name="example"></a>Exemple

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
