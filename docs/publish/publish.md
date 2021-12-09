---
title: Déployer et publier des compléments Office
description: Méthodes et options pour déployer votre complément Office à des fins de test ou de distribution auprès des utilisateurs.
ms.date: 12/07/2021
ms.localizationpriority: high
ms.openlocfilehash: 81c02a36becb9ef3244f7754dda44d064cdd9925
ms.sourcegitcommit: e392e7f78c9914d15c4c2538c00f115ee3d38a26
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/08/2021
ms.locfileid: "61331077"
---
# <a name="deploy-and-publish-office-add-ins"></a>Déployer et publier des compléments Office

Vous pouvez utiliser l’une des méthodes pour déployer votre complément Office à des fins de test ou de distribution auprès des utilisateurs.

|**Méthode**|**Use...**|
|:---------|:------------|
|[Chargement de version test](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)|Dans le cadre de votre processus de développement, pour tester votre complément exécuté sur Windows, iPad, Mac ou dans un navigateur. (Pas pour les compléments de production.)|
|[Partage réseau](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Dans le cadre de votre processus de développement, pour tester votre complément exécuté sur Windows après avoir publié le complément sur un serveur autre que l’hôte local. (ne convient pas pour les compléments de production ou pour les tests sur iPad, Mac ou le web.)|
|[AppSource](/office/dev/store/submit-to-appsource-via-partner-center)|Pour distribuer publiquement votre complément auprès des utilisateurs.|
|[Centre d’administration Microsoft 365](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)|Dans un environnement de cloud, utilisez cette méthode pour distribuer votre complément auprès des utilisateurs de votre organisation à l’aide du Centre d’administration Microsoft 365. Cela s’effectue à travers les [applications intégrées](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) ou le [déploiement centralisé](/microsoft-365/admin/manage/centralized-deployment-of-add-ins). |
|[Catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|Dans un environnement local, pour distribuer votre complément auprès des utilisateurs de votre organisation.|
|[Serveur Exchange](#outlook-add-in-deployment)|Dans un environnement local ou en ligne, pour distribuer des compléments Outlook à des utilisateurs.|

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="deployment-options-by-office-application-and-add-in-type"></a>Options de déploiement par l’application Office et le type de complément

Les options de déploiement disponibles dépendent de l’application Office que vous ciblez et du type de complément que vous créez.

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Options de déploiement pour les compléments Word, Excel et PowerPoint

| Point d’extension | Chargement d’une version test | Partage réseau | AppSource | Centre d’administration Microsoft 365 | Catalogue SharePoint\* |
|:----------------|:-----------:|:-------------:|:---------:|:--------------------------:|:--------------------:|
| Contenu         | X           | X             | X         | X                          | X                    |
| Volet Office       | X           | X             | X         | X                          | X                    |
| Commande         | X           | X             | X         | X                          |                      |

&#42; Les catalogues SharePoint ne prennent pas en charge Office sur Mac.

### <a name="deployment-options-for-outlook-add-ins"></a>Options de déploiement pour les compléments Outlook

| Point d’extension | Chargement d’une version test | AppSource | Serveur Exchange |
|:----------------|:-----------:|:---------:|:---------------:|
| Application de messagerie        | X           | X         | X               |
| Commande         | X           | X         | X               |

## <a name="production-deployment-methods"></a>Méthodes de déploiement de production

Les sections suivantes fournissent des informations supplémentaires sur les méthodes de déploiement les plus fréquemment utilisées pour distribuer des compléments Office de production aux utilisateurs au sein d’une organisation.

Si vous souhaitez plus d’informations sur l’acquisition, l’insertion et l’exécution des compléments par les utilisateurs finaux, consultez l’article relatif aux [premiers pas de l’utilisation de votre complément Office](https://support.microsoft.com/office/82e665c4-6700-4b56-a3f3-ef5441996862).

### <a name="integrated-apps-via-the-microsoft-365-admin-center"></a>Applications intégrées via le Centre d’administration Microsoft 365

Le centre d’administration Microsoft 365 permet aux administrateurs de déployer facilement des compléments Office auprès d’utilisateurs et de groupes au sein de leur organisation. Les compléments déployés via le centre d’administration sont disponibles pour les utilisateurs directement dans leurs applications Office, sans qu’aucune configuration client ne soit requise. Vous pouvez utiliser des applications intégrées pour déployer des compléments internes ainsi que des compléments fournis par des éditeurs de logiciels indépendants. Les applications intégrées montrent également les compléments administrateurs et d’autres applications regroupées par le même éditeur de logiciels indépendant, ce qui leur permet d’accéder à l’ensemble de l’expérience sur la plateforme Microsoft 365.

Lorsque vous liez vos compléments Office, applications Teams, applications SPFx et [d’autres applications](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps#what-apps-can-i-deploy-from-integrated-apps) ensemble, vous créez une offre SaaS (software as a service) unique pour vos clients. Pour obtenir des informations générales sur ce processus, consultez [Comment planifier une offre SaaS pour la place de marché commerciale](/azure/marketplace/plan-saas-offer). Pour plus d’informations sur la création d’applications intégrées, consultez [Configurer l’intégration de l’application Microsoft 365](/azure/marketplace/create-new-saas-offer#configure-microsoft-365-app-integration).

Pour plus d’informations sur le processus de déploiement des applications intégrées, consultez [Tester et déployer Microsoft 365 Apps par des partenaires dans le portail des applications intégrées](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps).

> [!IMPORTANT]
> Les clients des clouds souverains ou gouvernementaux n’ont pas accès aux applications intégrées. Ils utiliseront plutôt le déploiement centralisé. Le déploiement centralisé est une méthode de déploiement similaire, mais n’expose pas les compléments et applications connectés à l’administrateur. Pour plus d’informations, consultez [Déterminer si le déploiement centralisé des compléments fonctionne pour votre organisation](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

### <a name="sharepoint-app-catalog-deployment"></a>Déploiement d’un catalogue d’applications SharePoint

Un catalogue d’applications SharePoint est une collection de sites spéciale que vous pouvez créer pour héberger des compléments Word, Excel et PowerPoint. Les catalogues SharePoint ne prennent pas en charge les nouvelles fonctionnalités de complément mises en œuvre dans le nœud `VersionOverrides` du manifeste, y compris les commandes de complément. Nous vous recommandons d’utiliser Déploiement centralisé via le centre d’administration si possible. Par défaut, les commandes de complément déployées via un catalogue SharePoint s’ouvrent dans un volet des tâches.

Si vous déployez des compléments dans un environnement local, utilisez un catalogue SharePoint. Pour obtenir des détails, voir l’article sur la [publication de compléments du volet des tâches et de contenu dans un catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> Les catalogues SharePoint ne prennent pas en charge Office sur Mac. Pour déployer des compléments Office sur les clients Mac, vous devez les envoyer à [AppSource](/office/dev/store/submit-to-the-office-store).

### <a name="outlook-add-in-deployment"></a>Déploiement de compléments Outlook

Pour des environnements en ligne et locaux qui n’utilisent pas le service d’identité Azure AD, vous pouvez déployer des compléments Outlook via le serveur Exchange.

Le déploiement de compléments Outlook nécessite :

- Microsoft 365, Exchange Online ou Exchange Server 2013 ou version ultérieure
- Outlook 2013 ou une version ultérieure

Pour attribuer des compléments aux locataires, utilisez le centre d'administration Exchange pour télécharger un manifeste directement, à partir d'un fichier ou d'une URL, ou ajoutez un complément à partir d'AppSource. Pour affecter des compléments à des utilisateurs individuels, vous devez utiliser Exchange PowerShell. Pour plus de détails, consultez [Compléments pour Outlook dans Exchange Server](/exchange/add-ins-for-outlook-2013-help).

## <a name="see-also"></a>Voir aussi

- [Chargement de version test des compléments Outlook pour les tester](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Envoyer à AppSource][AppSource]
- [Instructions de conception pour les compléments Office](../design/add-in-design.md)
- [Création de descriptions efficaces dans AppSource](/office/dev/store/create-effective-office-store-listings)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
- [Qu’est-ce que la Place de marché commerciale Microsoft ?](/azure/marketplace/overview)

[AppSource]: /office/dev/store/submit-to-appsource-via-partner-center
