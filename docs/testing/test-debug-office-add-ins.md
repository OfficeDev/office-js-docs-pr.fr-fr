---
title: Test des compléments Office
description: Découvrez comment tester votre complément Office.
ms.date: 07/28/2022
ms.localizationpriority: high
ms.openlocfilehash: 56052182eafae59d42044ce4be40e086e51e8103
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467243"
---
# <a name="test-office-add-ins"></a>Test des compléments Office

Cet article contient des recommandations sur les tests, le débogage et la résolution des problèmes avec les compléments Office.

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a>Tester sur plusieurs plateformes et pour plusieurs versions d’Office

Les compléments Office s’exécutent sur les principales plateformes. Vous devez donc tester un complément sur toutes les plateformes sur lesquelles vos utilisateurs peuvent exécuter Office. Cela inclut généralement Office sur le Web, Office sur Windows (abonnement perpétuel et Microsoft 365), Office sur Mac, Office sur iOS et (pour les compléments Outlook) Office sur Android. Toutefois, dans certaines situations, vous pouvez être sûr qu’aucun de vos utilisateurs ne travaillera sur certaines plateformes. Par exemple, si vous créez un complément pour une entreprise qui exige que ses utilisateurs travaillent avec des ordinateurs Windows et un abonnement Office, vous n’avez pas besoin de tester Office sur Mac ou Office perpétuel sur Windows.

> [!NOTE]
> Sur les ordinateurs Windows, la version de Windows et d’Office détermine le contrôle de navigateur utilisé par les compléments. Pour plus d’informations, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Les compléments commercialisés via AppSource passent par un processus de validation qui inclut des tests sur toutes les plateformes. En outre, les compléments sont testés pour Office sur le web avec tous les principaux navigateurs modernes, y compris Microsoft Edge (WebView2 basé sur Chromium), Chrome et Safari. Par conséquent, vous devez effectuer des tests sur ces plateformes et navigateurs avant de les soumettre à AppSource. Pour plus d’informations sur la validation, consultez [Politiques de certification du marketplace commercial](/legal/marketplace/certification-policies), en particulier [section 1120.3](/legal/marketplace/certification-policies#11203-functionality)et la[Page de disponibilité et d’application de complément Office](/javascript/api/requirement-sets).
>
> AppSource n’utilise pas Internet Explorer ou la version héritée de Microsoft Edge (WebView1) pour tester les compléments dans Office sur le web. Toutefois, si un nombre important d’utilisateurs utiliseront Edge hérité pour ouvrir Office sur le web, vous devez le tester. (Office sur le web ne s’ouvre pas dans Internet Explorer, vous ne pouvez donc pas et n’avez pas besoin de tester Office sur le web avec Internet Explorer.) Pour plus d’informations, consultez [Support Internet Explorer 11](../develop/support-ie-11.md) et [Résolution des problèmes Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshoot-microsoft-edge-issues). Office prend toujours en charge ces navigateurs pour les runtimes de compléments. Par conséquent, si vous pensez avoir rencontré un bogue dans la façon dont les compléments s’exécutent dans ces derniers, créez un problème pour le dépôt [office-js.](https://github.com/OfficeDev/office-js/issues/new/choose)

## <a name="sideload-an-office-add-in-for-testing"></a>Chargement de version test d’un complément Office

You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product.

- [Chargement de version test des compléments Office sur Windows](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [Chargement de version test des compléments Office dans Office sur le web](sideload-office-add-ins-for-testing.md)

- [Chargement de versions test de compléments Office sur Mac](sideload-an-office-add-in-on-mac.md)

- [Chargement de versions test de compléments Office sur iPad](sideload-an-office-add-in-on-ipad.md)

- [Chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="unit-testing"></a>Tests unitaires

Pour plus d’informations sur l’ajout de tests unitaires à votre projet de complément, consultez [Test unitaires dans les compléments Office](unit-testing.md).

## <a name="debug-an-office-add-in"></a>Débogage d’un complément Office

La procédure de débogage d’un complément Office varie en fonction de votre plateforme et de votre environnement. Pour plus d’informations, consultez [Test et débogage de compléments Office](debug-add-ins-overview.md).

## <a name="validate-an-office-add-in-manifest"></a>Validation d’un manifeste de complément Office

Pour savoir comment valider le fichier manifeste qui décrit votre complément Office et résoudre des problèmes avec le fichier manifeste, consultez l’article [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md).

## <a name="troubleshoot-user-errors"></a>Résolution des erreurs de l’utilisateur

Pour plus d’informations sur la résolution des problèmes courants que les utilisateurs peuvent rencontrer avec votre complément Office, consultez [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md).
