---
title: Élément Requirements dans le fichier manifest
description: L’élément Requirements spécifie l’ensemble minimal de conditions requises et les méthodes dont votre Office a besoin pour s’activer.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 3a5a393485094b5cc830b5120c3abd8c211eff1e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153099"
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

|Élément|Contenu|Courrier|TaskPane|
|:-----|:-----|:-----|:-----|
|[Ensembles](sets.md)|x|x|x|
|[Méthodes](methods.md)|x||x|

## <a name="remarks"></a>Remarques

Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).
