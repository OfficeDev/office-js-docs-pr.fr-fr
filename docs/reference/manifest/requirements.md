---
title: Élément Requirements dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 364ab7c943895e1acecedba7970e54da331a2e6f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870365"
---
# <a name="requirements-element"></a>Élément Requirements

Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles des conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.

**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)

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

Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

