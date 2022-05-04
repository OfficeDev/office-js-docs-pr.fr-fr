---
title: Gérer votre complément Office
description: Comprendre nos engagements en matière de compatibilité et comment maintenir votre complément à jour.
ms.date: 04/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: 55da05d5c0b220adbeb0b4dbe248aa79f05b6b74
ms.sourcegitcommit: 5bf28c447c5b60e2cc7e7a2155db66cd9fe2ab6b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/04/2022
ms.locfileid: "65187354"
---
# <a name="maintain-your-office-add-in"></a>Gérer votre complément Office

Après avoir publié votre complément, vous devez le tenir à jour avec les modifications importantes apportées aux bibliothèques en amont. La mise à jour corrective des problèmes de sécurité est essentielle pour renforcer la confiance des clients. Étant donné que ces modifications n’ont aucun effet sur le manifeste publié, vos clients n’ont pas besoin d’effectuer d’actions pour obtenir les dernières versions de votre complément.

## <a name="breaking-changes-in-officejs"></a>Changements cassants dans Office.js

La plateforme de développement Microsoft 365 s’engage à garantir la compatibilité de votre complément. Nous nous efforçons d’éviter d’apporter des changements cassants à la surface et au comportement de l’API. Toutefois, dans certains cas, nous devons effectuer des mises à jour cassants pour des raisons de sécurité ou de fiabilité. Dans ces rares cas, les étapes suivantes sont prises pour garantir que les utilisateurs de votre complément ne sont pas affectés.

- Les annonces qui décrivent les fonctionnalités impactées et les modifications recommandées sont effectuées sur le [blog du développeur Microsoft 365](https://devblogs.microsoft.com/microsoft365dev/).
- Si votre complément est publié dans [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), vous serez contacté via les informations que vous avez fournies.
- Si possible, les administrateurs des locataires Microsoft 365 concernés (y compris [les locataires de développeurs) sont contactés](https://developer.microsoft.com/microsoft-365/dev-program) via le [Centre](/microsoft-365/admin/manage/message-center) de messages. Il incombe à l’administrateur de contacter les fournisseurs de solutions de complément publiées en dehors d’AppSource.

### <a name="app-assure"></a>Soutien aux applications

Le service [App Assure](https://www.microsoft.com/fasttrack/microsoft-365/app-assure) de Microsoft remplit la promesse de compatibilité des applications de Microsoft : vos applications fonctionneront sur Windows et Microsoft 365 Apps. Les ingénieurs App Assure sont disponibles pour vous aider à résoudre les problèmes que vous pouvez rencontrer sans coût supplémentaire.

Si vous rencontrez un problème de compatibilité d’application, les ingénieurs App Assure travailleront avec vous pour vous aider à résoudre le problème. Nos experts :

- Vous aider à résoudre les problèmes et à identifier une cause racine.
- Fournissez des conseils pour vous aider à résoudre le problème de compatibilité des applications.
- Contactez les éditeurs de logiciels indépendants (ISV) en votre nom pour corriger une partie de leur application, afin qu’elle soit fonctionnelle sur la version la plus moderne de nos produits.
- Collaborez avec les équipes d’ingénierie de produits Microsoft pour corriger les bogues de produit.

Pour en savoir plus sur App Assure, regardez [Apportez vos applications à Microsoft Edge avec App Assure : conseils et astuces](https://techcommunity.microsoft.com/t5/video-hub/bring-your-apps-to-microsoft-edge-with-app-assure-tips-and/ba-p/2167619). Pour envoyer votre demande de compatibilité des applications avec App Assure, remplissez le [formulaire d’inscription Microsoft FastTrack](https://aka.ms/AppAssureRequest) ou envoyez un e-mail à [achelp@microsoft.com](mailto:achelp@microsoft.com).

## <a name="changes-to-yeoman-templates-and-web-dependencies"></a>Modifications apportées aux modèles Yeoman et aux dépendances web

Le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md) s’appuie sur un certain nombre de bibliothèques de Microsoft et d’autres. Ces bibliothèques sont mises à jour indépendamment de toute activité Microsoft 365. Tous les projets créés avec le générateur doivent être tenus à jour à mesure que vous développez, publiez et gérez votre complément. Les outils suivants peuvent vous aider à vous assurer que votre projet utilise des versions sécurisées de toutes les bibliothèques dépendantes.

- [audit npm](https://docs.npmjs.com/cli/v6/commands/npm-audit/)
- [Dependabot et d’autres fonctionnalités de sécurité GitHub](https://github.com/features/security)

Ces instructions s’appliquent également aux copies d’exemples provenant des [Office exemples de code de complément](https://github.com/OfficeDev/Office-Add-in-samples) et d’autres sources.

### <a name="officejs-npm-package"></a>package NPM office.js

Le [package NPM office-js](https://www.npmjs.com/package/@microsoft/office-js) est une copie de ce qui est hébergé sur le [ réseau de distribution de contenuOffice.js (CDN).](../develop/understanding-the-javascript-api-for-office.md#accessing-the-office-javascript-api-library) Il est destiné aux scénarios où l’accès direct au CDN n’est pas possible. Le package NPM n’est pas destiné à fournir des références avec version à office.js. Nous vous recommandons vivement de toujours utiliser le CDN pour vous assurer que vous utilisez la dernière version des API JavaScript Office.

## <a name="current-best-practices"></a>Meilleures pratiques actuelles

Bien que nous nous efforçions de maintenir la compatibilité descendante, les modèles et les pratiques que nous recommandons évoluent continuellement. Notre documentation s’efforce de présenter les meilleures pratiques actuelles. Pour rester informé des nouvelles fonctionnalités susceptibles d’améliorer vos fonctionnalités existantes, rejoignez nos [compléments Office mensuels Community Call](../overview/office-add-ins-community-call.md).
