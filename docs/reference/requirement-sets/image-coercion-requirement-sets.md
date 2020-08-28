---
title: Ensembles de conditions requises de coercition d’image
description: Prise en charge des ensembles de conditions requises de forçage d’image avec des compléments Office dans Excel, PowerPoint et Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 7140099757c6e4b5ad405723d5fed95fded6d919
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293547"
---
# <a name="image-coercion-requirement-sets"></a>Ensembles de conditions requises de coercition d’image

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API dont un complément a besoin. Pour plus d’informations, consultez la rubrique [versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1,1 permet la conversion en image ( `Office.CoercionType.Image` ) lors de l’écriture de données à l’aide de la [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) méthode. Les applications suivantes sont prises en charge :

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

ImageCoercion 1,2 permet d’effectuer une conversion au format SVG ( `Office.CoercionType.XmlSvg` ) lors de l’écriture de données à l’aide de la [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) méthode. Les applications suivantes sont prises en charge :

- Excel sur Windows (connecté à un abonnement Microsoft 365)
- Excel sur Mac (connecté à un abonnement Microsoft 365)
- PowerPoint sur Windows (connecté à un abonnement Microsoft 365)
- PowerPoint sur Mac (connecté à un abonnement Microsoft 365)
- PowerPoint sur le web
- Word sur Windows (connecté à un abonnement Microsoft 365)
- Word sur Mac (connecté à un abonnement Microsoft 365)
- Word sur le web

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécification des exigences en matière d’applications et d’API Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
