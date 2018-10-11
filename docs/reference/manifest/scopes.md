# <a name="scopes-element"></a>Élément Scope

Contient des autorisations Microsoft Graph requises par le complément. L’Office Store se sert de l’élément Scope pour créer une boîte de dialogue de consentement. Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Type  |  Description  |
|:-----|:-----|:-----|
|  **Scope**                |  string     |   Nom d’une autorisation Microsoft Graph. Par exemple, Files.Read.All. |

## <a name="example"></a>Exemple

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
