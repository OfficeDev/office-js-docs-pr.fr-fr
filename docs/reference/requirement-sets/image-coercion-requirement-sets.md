---
title: Ensembles de conditions requises de coercition d’image
description: Prise en charge des ensembles de conditions requises pour le foragage d’image avec Office des Excel, PowerPoint et Word.
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 0f0b80c0af8213eaa9e3695373ddc037c2e60cc3
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/19/2021
ms.locfileid: "59450792"
---
# <a name="image-coercion-requirement-sets"></a>Ensembles de conditions requises de coercition d’image

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 permet la conversion en image ( ) lors de l’écriture de données `Office.CoercionType.Image` à l’aide de la [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) méthode. Les applications suivantes sont pris en charge.

- Excel 2013 et les ultérieures Windows
- Excel 2016 et ultérieures sur Mac
- Excel sur iPad
- OneNote sur le web
- PowerPoint 2013 et les Windows
- PowerPoint 2016 et ultérieures sur Mac
- PowerPoint sur le web
- PowerPoint sur iPad
- Word 2013 ou version ultérieure sur Windows
- Word 2016 ou version ultérieure sur Mac
- Word sur le web
- Word sur iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 permet la conversion au format SVG () lors de l’écriture de données `Office.CoercionType.XmlSvg` à l’aide de la [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) méthode. Les applications suivantes sont pris en charge.

- Excel 2021 et les Windows
- Excel 2021 et les ultérieures sur Mac
- PowerPoint 2021 et les Windows
- PowerPoint 2021 et les ultérieures sur Mac
- PowerPoint sur le web
- Word 2021 et les Windows
- Word 2021 et les ultérieurs sur Mac

## <a name="office-common-api-requirement-sets"></a>Séries de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
