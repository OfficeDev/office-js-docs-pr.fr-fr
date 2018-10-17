# <a name="desktopformfactor-element"></a>Élément DesktopFormFactor

Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau. Le facteur de forme pour bureau inclut Office pour Windows, Office pour Mac et Office Online. Il contient toutes les informations de complément pour ce facteur de forme à l’exception du nœud **Resources**.

Chaque définition de facteur de forme pour bureau contient l’élément **FunctionFile** et au moins un élément **ExtensionPoint**. Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).

## <a name="child-elements"></a>Éléments enfants

| Élément                               | Obligatoire | Description  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Oui      | Définit l’emplacement où se trouvent les fonctionnalités d’un complément |
| [FunctionFile](functionfile.md)       | Oui      | URL pointant vers un fichier qui contient les fonctions JavaScript.|
| [GetStarted](getstarted.md)           | Non       | Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint. |
| [SupportsSharedFolders](supportssharedfolders.md) | Non | Définit si le complément Outlook est disponible dans les scénarios délégués et est défini sur *false* par défaut.<br><br>**Important** : cet élément est disponible uniquement dans l'ensemble des conditions requises de la préversion des compléments Outlook sur Exchange Online. Les compléments qui utilisent cet élément ne peuvent pas être publiés sur AppSource ou déployés via un déploiement centralisé. |

## <a name="desktopformfactor-example"></a>Exemple pour DesktopFormFactor

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
