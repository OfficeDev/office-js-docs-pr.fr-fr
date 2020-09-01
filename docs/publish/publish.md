---
title: Déployer et publier des compléments Office
description: Méthodes et options pour déployer votre complément Office à des fins de test ou de distribution auprès des utilisateurs.
ms.date: 06/02/2020
localization_priority: Priority
ms.openlocfilehash: 7792b2196b686f36b81869ce81a3e3b87659f1c2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292393"
---
# <a name="deploy-and-publish-office-add-ins"></a>Déployer et publier des compléments Office

Vous pouvez utiliser l’une des méthodes pour déployer votre complément Office à des fins de test ou de distribution auprès des utilisateurs.

|**Méthode**|**Use...**|
|:---------|:------------|
|[Chargement de version test](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)|Dans le cadre de votre processus de développement, pour tester l’exécution de votre complément sur Windows, iPad, Mac ou dans un navigateur. (ne convient pas pour les compléments de production).|
|[Partage réseau](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Dans le cadre de votre processus de développement, pour tester votre complément exécuté sur Windows après avoir publié le complément sur un serveur autre que l’hôte local. (ne convient pas pour les compléments de production ou pour les tests sur iPad, Mac ou le web.)|
|[Déploiement centralisé](centralized-deployment.md)|Dans un environnement de cloud, utilisez cette méthode pour distribuer votre complément auprès des utilisateurs de votre organisation à l’aide du Centre d’administration Microsoft 365.|
|[Catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|Dans un environnement local, pour distribuer votre complément auprès des utilisateurs de votre organisation.|
|[AppSource](/office/dev/store/submit-to-appsource-via-partner-center)|Pour distribuer publiquement votre complément auprès des utilisateurs.|
|[Serveur Exchange](#outlook-add-in-deployment)|Dans un environnement local ou en ligne, pour distribuer des compléments Outlook à des utilisateurs.|

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="deployment-options-by-office-application-and-add-in-type"></a>Options de déploiement par l’application Office et le type de complément

Les options de déploiement disponibles dépendent de l’application Office que vous ciblez et du type de complément que vous créez.

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Options de déploiement pour les compléments Word, Excel et PowerPoint

| Point d’extension | Chargement d’une version test | Partage réseau | Centre d’administration Microsoft 365 |AppSource   | Catalogue SharePoint\* |
|:----------------|:-----------:|:-------------:|:-----------------------:|:----------:|:--------------------:|
| Contenu         | X           | X             | X                       | X          | X                    |
| Volet Office       | X           | X             | X                       | X          | X                    |
| Commande         | X           | X             | X                       | X          |                      |

&#42; Les catalogues SharePoint ne prennent pas en charge Office sur Mac.

### <a name="deployment-options-for-outlook-add-ins"></a>Options de déploiement pour les compléments Outlook

| Point d’extension | Chargement de version test | Serveur Exchange | AppSource    |
|:----------------|:-----------:|:---------------:|:------------:|
| Application de messagerie        | X           | X               | X            |
| Commande         | X           | X               | X            |

## <a name="production-deployment-methods"></a>Méthodes de déploiement de production

Les sections suivantes fournissent des informations supplémentaires sur les méthodes de déploiement les plus fréquemment utilisées pour distribuer des compléments Office de production aux utilisateurs au sein d’une organisation.

Si vous souhaitez plus d’informations sur l’acquisition, l’insertion et l’exécution des compléments par les utilisateurs finaux, consultez l’article relatif aux [premiers pas de l’utilisation de votre complément Office](https://support.office.com/article/start-using-your-office-add-in-82e665c4-6700-4b56-a3f3-ef5441996862).

### <a name="centralized-deployment-via-the-microsoft-365-admin-center"></a>Déploiement centralisé via le centre d’administration Microsoft 365

Le centre d’administration Microsoft 365 permet aux administrateurs de déployer facilement des compléments Office auprès d’utilisateurs et de groupes au sein de leur organisation. Les compléments déployés via le centre d’administration sont disponibles pour les utilisateurs directement dans leurs applications Office, sans qu’aucune configuration client ne soit requise. Vous pouvez utiliser le déploiement centralisé pour déployer des compléments internes, ainsi que des compléments fournis par des éditeurs de logiciels indépendants.

Pour plus d’informations, reportez-vous à [Publication des compléments Office à l’aide du déploiement centralisé via le centre d’administration Microsoft 365](centralized-deployment.md).

### <a name="sharepoint-app-catalog-deployment"></a>Déploiement d’un catalogue d’applications SharePoint

Un catalogue d’applications SharePoint est une collection de sites spéciale que vous pouvez créer pour héberger des compléments Word, Excel et PowerPoint. Les catalogues SharePoint ne prennent pas en charge les nouvelles fonctionnalités de complément mises en œuvre dans le nœud `VersionOverrides` du manifeste, y compris les commandes de complément. Nous vous recommandons d’utiliser Déploiement centralisé via le centre d’administration si possible. Par défaut, les commandes de complément déployées via un catalogue SharePoint s’ouvrent dans un volet des tâches.

Si vous déployez des compléments dans un environnement local, utilisez un catalogue SharePoint. Pour obtenir des détails, voir l’article sur la [publication de compléments du volet des tâches et de contenu dans un catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> Les catalogues SharePoint ne prennent pas en charge Office sur Mac. Pour déployer des compléments Office sur les clients Mac, vous devez les envoyer à [AppSource](/office/dev/store/submit-to-the-office-store).

### <a name="outlook-add-in-deployment"></a>Déploiement de compléments Outlook

Pour des environnements en ligne et locaux qui n’utilisent pas le service d’identité Azure AD, vous pouvez déployer des compléments Outlook via le serveur Exchange.

Le déploiement de compléments Outlook nécessite :

- Microsoft 365, Exchange Online ou Exchange Server 2013 ou version ultérieure
- Outlook 2013 ou une version ultérieure

Pour affecter des compléments à des clients, utilisez le centre d’administration Exchange pour télécharger un manifeste directement, à partir d’un fichier ou d’une URL, ou ajoutez un complément à partir d’AppSource. Pour affecter des compléments à des utilisateurs individuels, vous devez utiliser Exchange PowerShell. Pour plus d’informations, reportez-vous à [Installation ou suppression de compléments Outlook pour votre organisation](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx) sur TechNet.

## <a name="see-also"></a>Voir aussi

- [Chargement de version test des compléments Outlook pour les tester](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Envoyer à AppSource][AppSource]
- [Instructions de conception pour les compléments Office](../design/add-in-design.md)
- [Création de descriptions efficaces dans AppSource](/office/dev/store/create-effective-office-store-listings)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)

[AppSource]: /office/dev/store/submit-to-appsource-via-partner-center
