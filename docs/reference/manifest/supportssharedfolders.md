# <a name="supportssharedfolders-element"></a>Élément SupportsSharedFolders

Définit si le complément Outlook est disponible dans les scénarios de délégué. L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md). Il est défini sur *false* par défaut.

> [!IMPORTANT]
> Cet élément est disponible uniquement dans [l’ensemble des conditions requises de la préversion des compléments Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)  sur Exchange Online. Les compléments qui utilisent cet élément ne peuvent pas être publiés sur AppSource ou déployés via un déploiement centralisé.

Vous trouverez ci-dessous un exemple de l’élément **SupportsSharedFolders** .

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
