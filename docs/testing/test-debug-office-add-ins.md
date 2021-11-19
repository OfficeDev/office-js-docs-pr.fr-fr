---
title: Test et débogage de compléments Office
description: Découvrez comment tester et déboguer votre Complément Office.
ms.date: 09/24/2021
ms.localizationpriority: high
ms.openlocfilehash: db0edec5c7b7c741425a9d27d7580a52d2839546
ms.sourcegitcommit: 997a20f9fb011b96a50ceb04a4b9943d92d6ecf4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/19/2021
ms.locfileid: "61081413"
---
# <a name="test-and-debug-office-add-ins"></a>Test et débogage de compléments Office

Cet article contient des recommandations sur les tests, le débogage et la résolution des problèmes avec les compléments Office.

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a>Tester sur plusieurs plateformes et pour plusieurs versions d’Office

Les compléments Office s’exécutent sur les principales plateformes. Vous devez donc tester un complément sur toutes les plateformes sur lesquelles vos utilisateurs peuvent exécuter Office. Cela inclut généralement Office sur le web, Office sur Windows (abonnement et achat unique), Office sur Mac, Office sur iOS et (pour les compléments Outlook) Office sur Android. Toutefois, dans certaines situations, vous pouvez être sûr qu’aucun de vos utilisateurs ne travaillera sur certaines plateformes. Par exemple, si vous créez un complément pour une entreprise qui exige que ses utilisateurs utilisent des ordinateurs Windows et un abonnement Office, vous n’avez pas besoin de tester Office sur Mac ou Windows achat unique.

> [!NOTE]
> Sur les ordinateurs Windows, la version de Windows et d’Office détermine le contrôle de navigateur utilisé par les compléments. Pour plus d’informations, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Les compléments commercialisés via AppSource passent par un processus de validation qui inclut des tests sur toutes les plateformes. En outre, les compléments sont testés pour Office sur le web avec tous les principaux navigateurs modernes, y compris Microsoft Edge (WebView2 basé sur Chromium), Chrome et Safari. Par conséquent, vous devez effectuer des tests sur ces plateformes et navigateurs avant de les soumettre à AppSource. Pour plus d’informations sur la validation, consultez [Politiques de certification du marketplace commercial](/legal/marketplace/certification-policies), en particulier [section 1120.3](/legal/marketplace/certification-policies#11203-functionality)et la[Page de disponibilité et d’application de complément Office](../overview/office-add-in-availability.md).
>
> AppSource n’utilise pas Internet Explorer ou la version héritée de Microsoft Edge (WebView1) pour tester les compléments dans Office sur le web. Toutefois, si un nombre important d’utilisateurs utiliseront Edge hérité pour ouvrir Office sur le web, vous devez le tester. (Office sur le web ne s’ouvre pas dans Internet Explorer, vous ne pouvez donc pas et n’avez pas besoin de tester Office sur le web avec Internet Explorer.) Pour plus d’informations, consultez [Support Internet Explorer 11](../develop/support-ie-11.md) et [Résolution des problèmes Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues). Office prend toujours en charge ces navigateurs pour les runtimes de compléments. Par conséquent, si vous pensez avoir rencontré un bogue dans la façon dont les compléments s’exécutent dans ces derniers, créez un problème pour le dépôt [office-js.](https://github.com/OfficeDev/office-js/issues/new/choose)

## <a name="sideload-an-office-add-in-for-testing"></a>Chargement de version test d’un complément Office

Vous pouvez utiliser le chargement indépendant pour installer un complément Office à des fins de test sans avoir à le placer au préalable dans un catalogue de compléments. La procédure de chargement indépendant d’un complément varie selon la plateforme et, dans certains cas, le produit. Les articles suivants décrivent chacun comment charger une version test des compléments Office sur une plateforme spécifique ou dans un produit spécifique.

- [Chargement de version test des compléments Office sur Windows](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [Chargement de version test des compléments Office dans Office sur le web](sideload-office-add-ins-for-testing.md)

- [Chargement de version test de compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)

- [Chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a>Débogage d’un complément Office

La procédure de débogage d’un complément Office varie également selon la plateforme. Chacun des articles suivants décrit comment déboguer des compléments Office sur une plateforme spécifique.

- [Attacher un débogueur à partir du volet Office (sur Windows)](attach-debugger-from-task-pane.md)
- [Déboguer des compléments à l’aide des outils de développement pour Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Déboguer des compléments à l’aide des outils de développement pour la version héritée Edge](debug-add-ins-using-devtools-edge-legacy.md)
- [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
- [Débogage de compléments dans Office sur le web](debug-add-ins-in-office-online.md)
- [Déboguer des compléments Office sur un Mac](debug-office-add-ins-on-ipad-and-mac.md)
- [Complément Microsoft Office Extension de débogueur pour Visual Studio Code](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a>Validation d’un manifeste de complément Office

Pour savoir comment valider le fichier manifeste qui décrit votre complément Office et résoudre des problèmes avec le fichier manifeste, consultez l’article [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md).

## <a name="troubleshoot-user-errors"></a>Résolution des erreurs de l’utilisateur

Pour plus d’informations sur la résolution des problèmes courants que les utilisateurs peuvent rencontrer avec votre complément Office, consultez [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md).
