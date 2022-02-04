---
title: Élément MobileFormFactor dans le fichier manifest
description: L’élément MobileFormFactor spécifie les paramètres du facteur de forme mobile d’un module.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="mobileformfactor-element"></a>Élément MobileFormFactor

Spécifie les paramètres d’un complément pour le facteur de forme pour environnement mobile. Il contient toutes les informations de complément pour ce facteur de forme pour environnement mobile pour le nœud **Resources**.

**Type de complément :** messagerie

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

Chaque **définition MobileFormFactor** contient **l’élément FunctionFile** et un ou plusieurs **éléments ExtensionPoint** . Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).

L’élément **MobileFormFactor** est défini dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) le contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.

## <a name="child-elements"></a>Éléments enfants

| Élément                             | Requis | Description  |
|:------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | Oui      | Définit l’emplacement où se trouvent les fonctionnalités d’un complément |
| [FunctionFile](functionfile.md)     | Oui      | URL pointant vers un fichier qui contient les fonctions JavaScript.|

## <a name="mobileformfactor-example"></a>Exemple MobileFormFactor

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
