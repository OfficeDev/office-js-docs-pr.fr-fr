---
title: Maintenir votre complément Office
description: Comprendre nos engagements en matière de compatibilité et comment maintenir votre complément à jour.
ms.date: 05/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: c7f70eab252af516ab8dda591668d48392ce9f04
ms.sourcegitcommit: e63d8e32b25a9987f4a39b92a342a82b37a3404c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/17/2022
ms.locfileid: "65432189"
---
# <a name="maintain-your-office-add-in"></a>Maintenir votre complément Office

Après avoir publié votre complément, vous devez le tenir à jour avec les modifications importantes apportées aux bibliothèques en amont. La mise à jour corrective des problèmes de sécurité est essentielle pour renforcer la confiance des clients. Étant donné que ces modifications n’ont aucun effet sur le manifeste publié, vos clients n’ont pas besoin d’effectuer d’actions pour obtenir les dernières versions de votre complément.

## <a name="breaking-changes-in-officejs"></a>Changements cassants dans Office.js

La plateforme de développement Microsoft 365 s’engage à garantir la compatibilité de votre complément. Nous nous efforçons d’éviter d’apporter des changements cassants à la surface et au comportement de l’API. Toutefois, dans certains cas, nous devons effectuer des mises à jour cassants pour des raisons de sécurité ou de fiabilité. Dans ces rares cas, les étapes suivantes sont prises pour garantir que les utilisateurs de votre complément ne sont pas affectés.

- Les annonces qui décrivent les fonctionnalités impactées et les modifications recommandées sont effectuées sur le [blog du développeur Microsoft 365](https://devblogs.microsoft.com/microsoft365dev/).
- Si votre complément est publié dans [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), vous serez contacté via les informations que vous avez fournies.
- Si possible, les administrateurs des locataires Microsoft 365 concernés (y compris [les locataires de développeurs) sont contactés](https://developer.microsoft.com/microsoft-365/dev-program) via le [Centre](/microsoft-365/admin/manage/message-center) de messages. Il incombe à l’administrateur de contacter les fournisseurs de solutions de complément publiées en dehors d’AppSource.

### <a name="deprecation-policy"></a>Stratégie de dépréciation

Les API ou outils avec de meilleures alternatives peuvent être dépréciés. Microsoft fait tout son possible pour déclarer quelque chose comme déprécié au moins 24 mois avant de le mettre hors service. De même, pour les API individuelles généralement disponibles, Microsoft estime qu’une API est hors service au moins 24 mois avant de la supprimer de la version GA.

La dépréciation ne signifie pas nécessairement que la fonctionnalité ou l’API sera supprimée et inutilisable par les développeurs. Il indique qu’après la période de 24 mois, Microsoft ne prendra plus en charge l’API ou la fonctionnalité.

Lorsqu’une API est marquée comme obsolète, nous vous recommandons vivement de migrer vers la dernière version dès que possible. Dans certains cas, nous annoncerons que les nouvelles applications doivent commencer à utiliser les nouvelles API peu de temps après la dépréciation des API d’origine. Dans ces cas, seules les applications actives qui utilisent actuellement les API déconseillées peuvent continuer à les utiliser.

> [!IMPORTANT]
> La période de dépréciation de 24 mois est accélérée si l’attente de cette durée présente un risque de sécurité pour votre complément ou Microsoft.

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

## <a name="community-engagement"></a>Community engagement

À mesure que des mises à jour sont proposées pour la plateforme de développement Microsoft 365, nous sommes à l’écoute des commentaires. Veuillez signaler les préoccupations, les conséquences potentielles ou d’autres questions aux canaux [répertoriés dans Office compléments des ressources supplémentaires](../resources/resources-links-help.md).
