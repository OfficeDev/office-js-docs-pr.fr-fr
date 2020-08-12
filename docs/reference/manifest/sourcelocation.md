---
title: Élément SourceLocation dans le fichier manifeste
description: L’élément SourceLocation spécifie les emplacements des fichiers source pour votre complément Office.
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 447adb7df7d0c59305fe5046357959fcd7824735
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641402"
---
# <a name="sourcelocation-element"></a>Élément SourceLocation

Spécifie les emplacements des fichiers source pour votre complément Office sous la forme d’une URL de 1 à 2018 caractères. L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Contenu dans

- [DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)
- [FormSettings](formsettings.md) (compléments de messagerie)
- [ExtensionPoint](extensionpoint.md) (contextuel et LaunchEvent (aperçu) des compléments de messagerie)

## <a name="can-contain"></a>Peut contenir

[Override](override.md)

## <a name="attributes"></a>Attributs

|Attribut|Type|Requis|Description|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obligatoire|Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|
