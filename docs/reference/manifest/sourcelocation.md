---
title: Élément SourceLocation dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7544e2bae480b9431c8912533ea1b761132a355e
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451975"
---
# <a name="sourcelocation-element"></a>Élément SourceLocation

Spécifie les emplacements des fichiers source pour votre complément Office sous forme d’URL comprenant entre 1 et 2 018 caractères. L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Contenu dans

- [DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)
- [FormSettings](formsettings.md) (compléments de messagerie)
- [ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels)

## <a name="can-contain"></a>Peut contenir

[Override](override.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obligatoire|Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|
