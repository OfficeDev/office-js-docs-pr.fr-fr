---
title: Ensembles de conditions requises de coercition d’image
description: Prise en charge des ensembles de conditions requises de forçage d’image avec des compléments Office dans Excel, PowerPoint et Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: ccc65f3c38e8ddc4bea88d897e6abda73aa61e64
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717473"
---
# <a name="image-coercion-requirement-sets"></a>Ensembles de conditions requises de coercition d’image

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1,1 permet la conversion en image (`Office.CoercionType.Image`) lors de l’écriture de [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) données à l’aide de la méthode. Les hôtes suivants sont pris en charge :

- Excel 2013 et versions ultérieures sur Windows
- Excel 2016 et versions ultérieures sur Mac
- Excel sur iPad
- OneNote sur le web
- PowerPoint 2013 et versions ultérieures sur Windows
- PowerPoint 2016 et versions ultérieures sur Mac
- PowerPoint sur le web
- PowerPoint sur iPad
- Word 2013 ou version ultérieure sur Windows
- Word 2016 ou version ultérieure sur Mac
- Word sur le web
- Word sur iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1,2 permet d’effectuer une conversion au`Office.CoercionType.XmlSvg`format SVG () lors de [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) l’écriture de données à l’aide de la méthode. Les hôtes suivants sont pris en charge :

- Excel sur Windows (connecté à un abonnement Office 365)
- Excel sur Mac (connecté à un abonnement Office 365)
- PowerPoint sur Windows (connecté à un abonnement Office 365)
- PowerPoint sur Mac (connecté à un abonnement Office 365)
- PowerPoint sur le web
- Word sur Windows (connecté à un abonnement Office 365)
- Word sur Mac (connecté à un abonnement Office 365)
- Word sur le web

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécification des exigences en matière d’hôtes Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
