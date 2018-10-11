# <a name="sourcelocation-element"></a>Élément SourceLocation

Définit l’emplacement d’une ressource requise par les éléments Script ou Page utilisés par les fonctions personnalisées dans Excel.

## <a name="attributes"></a>Attributs

| **Attribut** | **Obligatoire** | **Description**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | Oui          | Nom d’une ressource d’URL définie dans la section &lt;Ressources&gt; du manifeste. |

## <a name="child-elements"></a>Éléments enfants

Aucun

## <a name="example"></a>Exemple

```xml
<SourceLocation resid="pageURL"/>
```