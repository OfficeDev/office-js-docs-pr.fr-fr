---
title: Élément SourceLocation dans le fichier manifeste
description: L’élément SourceLocation spécifie les emplacements de fichiers sources pour votre Office de recherche.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 6830a26cf192802c97c486511695b4ace35ac8263cfcd30ceaf71398f0d83a07
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095006"
---
# <a name="sourcelocation-element"></a>Élément SourceLocation

Spécifie les emplacements de fichiers sources de votre Office sous la mesure d’une URL de 1 à 2 018 caractères. L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Contenu dans

- [DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)
- [FormSettings](formsettings.md) (compléments de messagerie)
- [ExtensionPoint](extensionpoint.md) (modules de messagerie contextuels et LaunchEvent)

## <a name="can-contain"></a>Peut contenir

[Override](override.md)

## <a name="attributes"></a>Attributs

|Attribut|Type|Requis|Description|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obligatoire|Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|
