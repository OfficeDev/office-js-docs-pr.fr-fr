---
title: Élément defaultSettings dans le fichier manifeste
description: Spécifie l’emplacement de la source par défaut et d’autres paramètres par défaut pour votre complément de contenu ou de volet des tâches.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ace4f971d342f98d0aca5c21a7a48ceaf2563a2f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611581"
---
# <a name="defaultsettings-element"></a>Élément DefaultSettings

Spécifie l’emplacement de la source par défaut et d’autres paramètres par défaut pour votre complément de contenu ou de volet des tâches.

**Type de complément :** Application de contenu et de volet Office

## <a name="syntax"></a>Syntaxe

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Peut contenir

|**Élément**|**Content**|**Messagerie**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Remarques

L’emplacement source et les autres paramètres de l’élément **DefaultSettings** s’appliquent uniquement aux compléments de contenu et du volet Office. Pour les compléments de messagerie, vous spécifiez les emplacements par défaut des fichiers sources et d’autres paramètres par défaut dans l’élément [FormSettings](formsettings.md) .

