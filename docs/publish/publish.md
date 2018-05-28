---
title: D?ploiement et publication de votre compl?ment Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: d8264667306dcdac2e9d5e5d6e6607a2a2100546
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="deploy-and-publish-your-office-add-in"></a>D?ploiement et publication de votre compl?ment Office

Vous pouvez utiliser l?une des m?thodes pour d?ployer votre compl?ment Office ? des fins de test ou de distribution aupr?s des utilisateurs.

|**M?thode**|**Use...**|
|:---------|:------------|
|[Chargement de version test](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Dans le cadre de votre processus de d?veloppement, pour tester l?ex?cution de votre compl?ment sur Windows, Office Online, iPad ou Mac.|
|[D?ploiement centralis?](centralized-deployment.md)|Dans un environnement de cloud ou hybride, utilisez cette m?thode pour distribuer votre compl?ment aupr?s des utilisateurs de votre organisation ? l?aide du Centre d?administration Office 365.|
|[Catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|Dans un environnement local, pour distribuer votre compl?ment aupr?s des utilisateurs de votre organisation.|
|[AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store)|Pour distribuer publiquement votre compl?ment aupr?s des utilisateurs.|
|[Serveur Exchange](#outlook-add-in-deployment)|Dans un environnement local ou en ligne, pour distribuer des compl?ments Outlook ? des utilisateurs.|
|[Partage r?seau](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Sur un ordinateur Windows sur un r?seau sur lequel vous voulez h?berger votre compl?ment, acc?dez au dossier parent ou ? la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partag?s.|

> [!NOTE]
> Si vous pr?voyez de [publier](../publish/publish.md) votre compl?ment sur AppSource et de le rendre disponible dans l?exp?rience Office, assurez-vous que vous respectez les [strat?gies de validation AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Par exemple, pour r?ussir la validation, votre compl?ment doit fonctionner sur toutes les plateformes prenant en charge les m?thodes d?finies (pour en savoir plus, consultez la [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative ? la disponibilit? des compl?ments Office sur les plateformes et les h?tes](../overview/office-add-in-availability.md)).

## <a name="deployment-options-by-office-host"></a>Options de d?ploiement par l?h?te Office

Les options de d?ploiement disponibles d?pendent de l?h?te Office que vous ciblez et du type de compl?ment que vous cr?ez.

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Options de d?ploiement pour les compl?ments Word, Excel et PowerPoint

| Point d?extension | Chargement de version test | Centre d?administration Office 365 |AppSource| Catalogue SharePoint\*  |
|:----------------|:-----------:|:-----------------------:|:----------:|:--------------------:|
| Contenu         | X           | X                       | X          | X                    |
| Volet Office       | X           | X                       | X          | X                    |
| Commande           | X           | X                       | X          |                      |

* Les catalogues SharePoint ne prennent pas en charge Office 2016 pour Mac.

### <a name="deployment-options-for-outlook-add-ins"></a>Options de d?ploiement pour les compl?ments Outlook

| Point d?extension | Chargement de version test | Serveur Exchange | AppSource |
|:----------------|:-----------:|:---------------:|:------------:|
| Application de messagerie        | X           | X               | X            |
| Commande         | X           | X               | X            |

## <a name="deployment-methods"></a>M?thodes de d?ploiement

Les sections suivantes fournissent des informations suppl?mentaires sur les m?thodes de d?ploiement les plus fr?quemment utilis?es pour distribuer des compl?ments Office aux utilisateurs au sein d?une organisation.

Pour plus d?informations sur l?acquisition, l?insertion et l?ex?cution des compl?ments par les utilisateurs finaux, consultez l?article relatif aux [premiers pas de l?utilisation de votre compl?ment Office](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

### <a name="centralized-deployment-via-the-office-365-admin-center"></a>D?ploiement centralis? via le centre d?administration Office 365 

Le centre d?administration Office 365 permet aux administrateurs de d?ployer facilement des compl?ments Office aupr?s d?utilisateurs et de groupes au sein de leur organisation. Les compl?ments d?ploy?s via le centre d?administration sont disponibles pour les utilisateurs directement dans leurs applications Office, sans qu?aucune configuration client ne soit requise. Vous pouvez utiliser le d?ploiement centralis? pour d?ployer des compl?ments internes, ainsi que des compl?ments fournis par des ?diteurs de logiciels ind?pendants.

Pour plus d?informations, reportez-vous ? [Publication des compl?ments Office ? l?aide du d?ploiement centralis? via le centre d?administration Office 365](centralized-deployment.md).

### <a name="sharepoint-catalog-deployment"></a>D?ploiement d?un catalogue SharePoint

Un catalogue de compl?ments SharePoint est une collection de sites sp?ciale que vous pouvez cr?er pour h?berger des compl?ments Word, Excel et PowerPoint. Les catalogues SharePoint ne prennent pas en charge les nouvelles fonctionnalit?s de compl?ment mises en ?uvre dans le n?ud `VersionOverrides` du manifeste, y compris les commandes de compl?ment. Nous vous recommandons d?utiliser D?ploiement centralis? via le centre d?administration si possible. Par d?faut, les commandes de compl?ment d?ploy?es via un catalogue SharePoint s?ouvrent dans un volet des t?ches.

Si vous d?ployez des compl?ments dans un environnement local, utilisez un catalogue SharePoint. Pour obtenir des d?tails, voir l?article sur la [publication de compl?ments du volet des t?ches et de contenu dans un catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> Les catalogues SharePoint ne prennent pas en charge Office 2016 pour Mac. Pour d?ployer des compl?ments Office sur les clients Mac, vous devez les envoyer ? [AppSource]. 

### <a name="outlook-add-in-deployment"></a>D?ploiement de compl?ments Outlook

Pour des environnements en ligne et locaux qui n?utilisent pas le service d?identit? Azure AD, vous pouvez d?ployer des compl?ments Outlook via le serveur Exchange. 

Le d?ploiement de compl?ments Outlook n?cessite :

- Office 365, Exchange Online ou Exchange Server 2013 ou version ult?rieure
- Outlook 2013 ou une version ult?rieure

Pour affecter des compl?ments ? des clients, utilisez le centre d?administration Exchange pour t?l?charger un manifeste directement, ? partir d?un fichier ou d?une URL, ou ajoutez un compl?ment ? partir d?AppSource. Pour affecter des compl?ments ? des utilisateurs individuels, vous devez utiliser Exchange PowerShell. Pour plus d?informations, reportez-vous ? [Installation ou suppression de compl?ments Outlook pour votre organisation](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx) sur TechNet.

## <a name="see-also"></a>Voir aussi

- [Chargement de version test des compl?ments Outlook](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Envoyer ? AppSource][AppSource]
- [Instructions de conception pour les compl?ments Office](../design/add-in-design.md)
- [Cr?ation de descriptions efficaces dans AppSource](https://docs.microsoft.com/en-us/office/dev/store/create-effective-office-store-listings)
- [R?solution des erreurs rencontr?es par l?utilisateur avec des compl?ments Office](../testing/testing-and-troubleshooting.md)

[AppSource]: https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store
[Office Add-in host and platform availability]: ../overview/office-add-in-availability
