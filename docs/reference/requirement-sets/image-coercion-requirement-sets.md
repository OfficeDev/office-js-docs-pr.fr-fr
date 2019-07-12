---
title: Ensembles de conditions requises de forçage d’image
description: Prise en charge des ensembles de conditions requises de forçage d’image avec des compléments Office dans Excel, PowerPoint et Word.
ms.date: 07/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 046a3f1f16d8b48cddbd64bddf80a31ed1e50583
ms.sourcegitcommit: 61f8f02193ce05da957418d938f0d94cb12c468d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/11/2019
ms.locfileid: "35633990"
---
# <a name="image-coercion-requirement-sets"></a>Ensembles de conditions requises de forçage d’image

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de forçage d’image, les applications hôtes Office qui prennent en charge l’ensemble de conditions requises, ainsi que les numéros de build ou de version de l’application Office.

## <a name="imagecoercion-11"></a>ImageCoercion 1,1

ImageCoercion 1,1 permet la conversion en image (`Office.CoercionType.Image`) lors de l’écriture de [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) données à l’aide de la méthode. Les hôtes suivants sont pris en charge:

- Excel 2013 et versions ultérieures sur Windows
- Excel 2016 et versions ultérieures sur Mac
- Excel sur le Web
- Excel sur iPad
- OneNote sur le Web
- PowerPoint 2013 et versions ultérieures sur Windows
- PowerPoint 2016 et versions ultérieures sur Mac
- PowerPoint sur le Web
- PowerPoint sur iPad
- Word 2013 et versions ultérieures sur Windows
- Word 2016 et versions ultérieures sur Mac
- Word sur le Web
- Word pour iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1,2

ImageCoercion 1,2 permet d’effectuer une conversion au`Office.CoercionType.XmlSvg`format SVG () lors de [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) l’écriture de données à l’aide de la méthode. Les hôtes suivants sont pris en charge:

- Excel sur Windows (connecté à un abonnement Office 365)
- Excel sur Mac (connecté à un abonnement Office 365)
- Excel sur le Web
- PowerPoint sur Windows (connecté à un abonnement Office 365)
- PowerPoint sur Mac (connecté à un abonnement Office 365)
- PowerPoint sur le Web
- Word sur Windows (connecté à un abonnement Office 365)
- Word sur Mac (connecté à un abonnement Office 365)
- Word sur le Web

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
