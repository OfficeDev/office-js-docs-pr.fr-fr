---
title: Élément Requirements dans le fichier manifest
description: L’élément Requirements spécifie l’ensemble de conditions requises minimum et les méthodes nécessaires à l’activation de votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 586f05ec68257462cb64a96abf2a34eb31861a5c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611714"
---
# <a name="requirements-element"></a>Élément Requirements

Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Peut contenir

|**Élément**|**Content**|**Messagerie**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Ensembles](sets.md)|x|x|x|
|[Méthodes](methods.md)|x||x|

## <a name="remarks"></a>Remarques

Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).
