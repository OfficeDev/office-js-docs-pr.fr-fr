---
title: Élément Requirements dans le fichier manifest
description: L’élément Requirements spécifie l’ensemble minimal de conditions requises et les méthodes dont votre Office a besoin pour s’activer.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3020037b48e3f759acf6a7e2758bb8c1fd2dd36429e0b21613e22fca33cacc1a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098104"
---
# <a name="requirements-element"></a>Élément Requirements

Spécifie l’ensemble minimal d’Office de l’API JavaScript[(ensembles](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) de conditions requises et/ou méthodes) que votre Office doit activer.

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

|Élément|Contenu|Courrier Outlook|TaskPane|
|:-----|:-----|:-----|:-----|
|[Ensembles](sets.md)|x|x|x|
|[Méthodes](methods.md)|x||x|

## <a name="remarks"></a>Remarques

Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).
