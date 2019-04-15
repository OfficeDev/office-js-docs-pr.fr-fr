---
title: Configuration requise pour exécuter des compléments Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: c57534a8d00904336af518d9d32606373b2edab6
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872038"
---
# <a name="requirements-for-running-office-add-ins"></a>Configuration requise pour exécuter des compléments Office

Cet article décrit la configuration logicielle et matérielle requise pour l’exécution des compléments Office.

> [!NOTE]
> Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de validation AppSource](/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)).

Pour savoir de manière détaillée quelle version d’Office prend en charge les compléments Office, consultez la page relative à la [disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md).

## <a name="server-requirements"></a>Exigences en matière de serveur

Pour pouvoir installer et exécuter des Complément Office, vous devez d’abord déployer les fichiers manifeste et de pages web pour l’interface utilisateur et le code de votre complément sur les emplacements de serveur appropriés.

Pour tous les types de complément (compléments de contenu, Outlook et volet Office, et les commandes de compléments), vous devez déployer les fichiers de pages web de votre complément sur un serveur web ou un service d’hébergement web, tel que [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> lorsque vous développez et déboguez un complément dans Visual Studio, Visual Studio déploie et exécute les fichiers de page web de votre complément localement avec IIS Express et ne nécessite aucun serveur web supplémentaires. 

Pour les compléments du volet Office et de contenu, dans les applications hôtes Office prises en charge (applications web Access, Word, Excel, PowerPoint ou Project), vous avez également besoin d’un [catalogue de compléments](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) sur SharePoint pour télécharger les fichiers manifeste XML des compléments.

Pour tester et exécuter des compléments Outlook, le compte de messagerie Outlook de l’utilisateur doit être situé sur Exchange 2013 ou une version ultérieure, disponible par le biais d’Office 365, Exchange Online ou via une installation sur site. L’utilisateur ou l’administrateur installe les fichiers manifeste pour les compléments Outlook sur ce serveur.

> [!NOTE]
> Les comptes de messagerie POP et IMAP dans Outlook ne prennent pas en charge les compléments Office.

## <a name="client-requirements-windows-desktop-and-tablet"></a>Exigences en matière de client : ordinateur de bureau et tablette Windows

Le logiciel suivant est requis pour développer un Complément Office pour les clients Office ou les clients web pris en charge qui s’exécutent sur un ordinateur de bureau, un ordinateur portable ou une tablette Windows :


- Pour les ordinateurs de bureau Windows x86 et x64 et les tablettes telles que Surface Pro :
    - La version 32 bits ou 64 bits d’Office 2013 ou une version ultérieure s’exécutant sur Windows 7 ou une version ultérieure.
    - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professionnel 2013, Project 2013 SP1 ou Word 2013, ou une version ultérieure du client Office, si vous testez ou exécutez un Complément Office, notamment pour l’un de ces clients de bureau Office. Les clients de bureau Office peuvent être installés sur site ou par le biais de « Démarrer en un clic » sur l’ordinateur client.

  Si votre abonnement Office 365 est valide, mais que n’avez pas accès au client Office, nous vous conseillons de [télécharger et installer la dernière version d’Office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658).

- Internet Explorer 11 ou version ultérieure, qui doit être installé, mais ne doit pas être le navigateur par défaut. Pour prendre en charge les compléments Office, le client Office qui s’exécute en tant qu’hôte utilise des composants de navigateur qui font partie d’Internet Explorer 11 ou version ultérieure.

  > [!NOTE]
  > La Configuration de sécurité renforcée d’Internet Explorer (ESC) doit être désactivée pour que les compléments web Office fonctionnent. Si vous utilisez un ordinateur Windows Server comme votre client lors du développement des compléments, notez qu’ESC est activée par défaut dans Windows Server.

- L’un des éléments suivants en tant que navigateur par défaut : Internet Explorer 11 ou version ultérieure, ou la dernière version de Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).
- Un éditeur HTML et JavaScript tel que le Bloc-notes, [Visual Studio et les outils de développement Office ](https://www.visualstudio.com/features/office-tools-vs) ou un outil de développement web tiers.

## <a name="client-requirements-os-x-desktop"></a>Exigences en matière de client : ordinateur de bureau OS X

Outlook pour Mac, qui est distribué dans le cadre d’Office 365, prend en charge les compléments Outlook. L’exécution des compléments Outlook sur Outlook pour Mac a les mêmes exigences qu’Outlook pour Mac lui-même : le système d’exploitation doit être au minimum OS X v10.10 « Yosemite ». Comme Outlook pour Mac utilise WebKit comme moteur de disposition pour restituer les pages de complément, il n’existe pas de dépendance de navigateur supplémentaire.

Les versions de client minimales d’Office pour Mac prenant en charge les compléments Office sont les suivantes :

- Word pour Mac version 15.18 (160109)
- Excel pour Mac version 15.19 (160206)
- PowerPoint pour Mac version 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-online-web-clients-and-sharepoint"></a>Exigences en matière de client : prise en charge du navigateur pour les clients web Office Online et SharePoint

Tout navigateur qui prend en charge ECMAScript 5.1, HTML5 et CSS3, tel qu’Internet Explorer 11 ou version ultérieure, ou la dernière version de Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Exigences en matière de client : smartphones et tablettes autres que Windows

L’installation du logiciel suivant est nécessaire pour tester et exécuter les compléments Outlook. Ce logiciel est conçu spécialement pour les versions d’Outlook Web App exécutées dans des navigateurs sur smartphones et tablettes utilisant d’autres systèmes d’exploitation que Windows.


| Application hôte | Appareil | Système d’exploitation | Compte Exchange | Navigateur mobile |
|:-----|:-----|:-----|:-----|:-----|
|Outlook pour Android|Tablettes et smartphones Android|Android KitKat 4.4 et version ultérieure|Sur la dernière mise à jour d’Office 365 pour les entreprises ou d’Exchange Online|Application native pour Android, navigateur non applicable|
|Outlook pour iOS|Tablettes iPad, smartphones iPhone|iOS 11 ou version ultérieure|Sur la dernière mise à jour d’Office 365 pour les entreprises ou d’Exchange Online|Application native pour iOS, navigateur non applicable|
|Outlook Web App|iPhone 4, iPad 2, iPod Touch 4 (ou version ultérieure de ces appareils)|iOS 5 ou version ultérieure|Sur Office 365, Exchange Online ou en local sur Exchange Server 2013 ou version ultérieure|Safari|

> [!NOTE]
> Les applications natives OWA pour Android, OWA pour iPad et OWA pour iPhone ont été [supprimées](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b) et ne sont plus requises ou disponibles pour les tests des compléments Outlook.


## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)
