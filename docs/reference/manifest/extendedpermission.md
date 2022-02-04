---
title: Élément ExtendedPermission dans le fichier manifeste
description: Définit une autorisation étendue dont le add-in a besoin pour accéder à l’API ou à la fonctionnalité associée.
ms.date: 01/04/2022
ms.localizationpriority: medium
---

# <a name="extendedpermission-element"></a>`ExtendedPermission` élément

Définit une autorisation étendue dont le add-in a besoin pour accéder à l’API ou à la fonctionnalité associée. L’élément `ExtendedPermission` est un élément enfant [de ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1.9. Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

**Type de complément :** messagerie

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [Mailbox 1.9](../../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)

## <a name="available-extended-permissions"></a>Autorisations étendues disponibles

Voici les valeurs disponibles.

|Valeur disponible|Description|Hôtes|
|---|---|---|
|`AppendOnSend`|Déclare que le add-in utilise le [Office. API Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-appendonsendasync-member(1)).|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` exemple

Voici un exemple de l’élément `ExtendedPermission` .

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="contained-in"></a>Contenu dans

[ExtendedPermissions](extendedpermissions.md)
